"""pdf checks."""

from __future__ import annotations

from typing import Any

from ..context import Context
from ..registry import (BOOL, COUNT_PROPS, COUNT_REQ, NONEMPTY_STR, POS_INT, TEXT_MATCH_PROPS, TEXT_OR_REGEX,
                        CheckError, Outcome, check)
from ..util import TextMatcher, count_expected, count_ok, first_n, norm_ws

PAGE = {"page": POS_INT}
# Portrait (width, height) in points.
PAGE_SIZES = {
    "A3": (841.89, 1190.55), "A4": (595.28, 841.89), "A5": (419.53, 595.28),
    "LETTER": (612.0, 792.0), "LEGAL": (612.0, 1008.0), "TABLOID": (792.0, 1224.0),
}


def _page_index(ctx: Context, page: int) -> int:
    n = len(ctx.pypdf().pages)
    if page > n:
        raise CheckError(f"document has {n} page(s), no page {page}")
    return page - 1


@check("pdf.page_count", COUNT_PROPS, **COUNT_REQ)
def page_count(ctx: Context, p: dict[str, Any]) -> Outcome:
    n = len(ctx.pypdf().pages)
    return Outcome(count_ok(n, p), count_expected(p), n)


@check("pdf.contains_text", {**PAGE, **TEXT_MATCH_PROPS}, **TEXT_OR_REGEX)
def contains_text(ctx: Context, p: dict[str, Any]) -> Outcome:
    m = TextMatcher(p, "contains")
    texts = ctx.pdf_page_texts()
    if "page" in p:
        i = _page_index(ctx, p["page"])
        found = [i + 1] if m.matches(texts[i]) else []
    else:
        found = [i + 1 for i, t in enumerate(texts) if m.matches(t)]
    exp = m.describe() | {"page": p.get("page", "any")}
    return Outcome(bool(found), exp, {"found_on_pages": found})


def _norm_field_value(v: Any) -> Any:
    if v is None:
        return None
    if isinstance(v, list):
        return [_norm_field_value(x) for x in v]
    s = str(v)
    return s[1:] if s.startswith("/") else s  # NameObject (/Yes, /Off) -> Yes, Off


@check("pdf.form_field_equals", {"name": NONEMPTY_STR,
                                 "value": {"type": ["string", "number", "boolean", "null", "array"]},
                                 "ignore_case": BOOL}, required=("name", "value"))
def form_field_equals(ctx: Context, p: dict[str, Any]) -> Outcome:
    fields = ctx.pypdf().get_fields() or {}
    name = p["name"]
    if name not in fields:
        short = [k for k in fields if k.rsplit(".", 1)[-1] == name]
        if len(short) != 1:
            return Outcome(False, {"name": name, "value": p["value"]},
                           {"error": "field not found", "fields": sorted(fields)[:50]})
        name = short[0]
    raw = fields[name].get("/V")
    actual = _norm_field_value(raw.get_object() if hasattr(raw, "get_object") else raw)
    want = p["value"]
    if isinstance(want, bool):
        on = actual not in (None, "", "Off", "off", "0", "false", "False")
        ok = on == want
    elif isinstance(want, (int, float)):
        try:
            ok = actual is not None and float(str(actual).replace(",", "")) == float(want)
        except ValueError:
            ok = False
    elif isinstance(want, list):
        got = actual if isinstance(actual, list) else ([actual] if actual is not None else [])
        ok = [str(x) for x in got] == [str(x) for x in want]
    elif want is None:
        ok = actual in (None, "")
    else:
        a = "" if actual is None else str(actual)
        ok = a.casefold() == want.casefold() if p.get("ignore_case") else a == want
    return Outcome(ok, {"name": name, "value": want}, actual)


@check("pdf.is_text_searchable", {"min_chars": POS_INT,
                                  "min_ratio": {"type": "number", "exclusiveMinimum": 0, "maximum": 1}})
def is_text_searchable(ctx: Context, p: dict[str, Any]) -> Outcome:
    texts = ctx.pdf_page_texts()
    need = p.get("min_chars", 1)
    ok_pages = [i + 1 for i, t in enumerate(texts) if len("".join(t.split())) >= need]
    ratio = len(ok_pages) / len(texts) if texts else 0.0
    want = p.get("min_ratio", 1.0)
    return Outcome(bool(texts) and ratio >= want, {"min_ratio": want, "min_chars": need},
                   {"ratio": round(ratio, 4), "pages_without_text": [i for i in range(1, len(texts) + 1)
                                                                     if i not in ok_pages][:50]})


def _page_dims(page: Any) -> tuple[float, float]:
    box = page.cropbox
    w, h = abs(float(box.width)), abs(float(box.height))
    if (page.rotation or 0) % 180 == 90:
        w, h = h, w
    return w, h


@check("pdf.page_size", {**PAGE, "size": {"enum": sorted(PAGE_SIZES) + [s.title() for s in PAGE_SIZES]},
                         "width": {"type": "number", "exclusiveMinimum": 0},
                         "height": {"type": "number", "exclusiveMinimum": 0},
                         "orientation": {"enum": ["portrait", "landscape"]},
                         "tolerance": {"type": "number", "minimum": 0}},
       anyOf=[{"required": ["size"]}, {"required": ["width", "height"]}, {"required": ["orientation"]}],
       dependentRequired={"width": ["height"], "height": ["width"]},
       **{"not": {"required": ["size", "width"]}})
def page_size(ctx: Context, p: dict[str, Any]) -> Outcome:
    reader = ctx.pypdf()
    pages = [_page_index(ctx, p["page"])] if "page" in p else range(len(reader.pages))
    tol = p.get("tolerance", 2.0)
    target = PAGE_SIZES[p["size"].upper()] if "size" in p else (
        (p["width"], p["height"]) if "width" in p else None)
    orient = p.get("orientation")
    bad, dims = [], []
    for i in pages:
        w, h = _page_dims(reader.pages[i])
        dims.append({"page": i + 1, "width": round(w, 2), "height": round(h, 2)})
        ok = True
        if target is not None:
            tw, th = target
            exact = abs(w - tw) <= tol and abs(h - th) <= tol
            swapped = abs(w - th) <= tol and abs(h - tw) <= tol
            ok = exact or (swapped and "size" in p)  # named sizes accept either orientation
        if orient is not None:
            ok = ok and (("landscape" if w > h else "portrait") == orient)
        if not ok:
            bad.append(i + 1)
    exp = {k: p[k] for k in ("page", "size", "width", "height", "orientation") if k in p} | {"tolerance": tol}
    return Outcome(not bad and bool(dims), exp, {"mismatched_pages": bad, "pages": first_n(dims, 20)})


def _outline_titles(outline: Any) -> list[str]:
    out = []
    for item in outline:
        if isinstance(item, list):
            out.extend(_outline_titles(item))
        else:
            out.append(norm_ws(str(getattr(item, "title", "") or "")))
    return out


@check("pdf.bookmarks_present", {"min_count": POS_INT, "titles": {"type": "array", "items": NONEMPTY_STR},
                                 "ignore_case": BOOL})
def bookmarks_present(ctx: Context, p: dict[str, Any]) -> Outcome:
    titles = _outline_titles(ctx.pypdf().outline)
    fold = (lambda s: s.casefold()) if p.get("ignore_case") else (lambda s: s)
    have = {fold(t) for t in titles}
    missing = [t for t in p.get("titles", []) if fold(norm_ws(t)) not in have]
    need = p.get("min_count", 1)
    exp: dict[str, Any] = {"min_count": need}
    if "titles" in p:
        exp["titles"] = p["titles"]
    act: dict[str, Any] = {"count": len(titles), "titles": first_n(titles, 50)}
    if missing:
        act["missing"] = missing
    return Outcome(len(titles) >= need and not missing, exp, act)


@check("pdf.table_on_page", {**PAGE, "min_rows": POS_INT, "min_cols": POS_INT,
                             "strategy": {"enum": ["lines", "text"]}})
def table_on_page(ctx: Context, p: dict[str, Any]) -> Outcome:
    pdf = ctx.plumber()
    idx = [_page_index(ctx, p["page"])] if "page" in p else range(len(pdf.pages))
    strategy = p.get("strategy", "lines")
    settings = {"vertical_strategy": strategy, "horizontal_strategy": strategy}
    found = []
    for i in idx:
        for t in pdf.pages[i].find_tables(table_settings=settings):
            rows = t.extract()
            rows = [r for r in rows if any(c not in (None, "") for c in r)]
            found.append({"page": i + 1, "rows": len(rows), "cols": max((len(r) for r in rows), default=0)})
    need_r, need_c = p.get("min_rows", 2), p.get("min_cols", 1)
    ok = any(t["rows"] >= need_r and t["cols"] >= need_c for t in found)
    exp = {"page": p.get("page", "any"), "min_rows": need_r, "min_cols": need_c, "strategy": strategy}
    return Outcome(ok, exp, {"tables": first_n(found, 20)})
