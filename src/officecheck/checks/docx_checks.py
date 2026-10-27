"""docx checks."""

from __future__ import annotations

import re
from typing import Any

from docx.enum.style import WD_STYLE_TYPE

from ..context import Context
from ..ooxml import W_P, field_instructions, parts_matching, q, read_xml, w_paragraph_texts, w_text
from ..registry import (BOOL, COUNT_PROPS, COUNT_REQ, NONEMPTY_STR, POS_INT,
                        STR_OR_LIST, TEXT_MATCH_PROPS, TEXT_OR_REGEX, CheckError, Outcome, check)
from ..util import TextMatcher, count_expected, count_ok, first_n, norm_ws

LEVEL = {"type": "integer", "minimum": 0, "maximum": 9}
_HEADING_NAME = re.compile(r"heading\s*([1-9])", re.IGNORECASE)


# ---- headings --------------------------------------------------------------------------------


def _style_level(style: Any) -> int | None:
    """Heading level implied by a style or its base styles (Title = 0)."""
    seen = 0
    while style is not None and seen < 20:
        for name in (style.name or "", style.style_id or ""):
            m = _HEADING_NAME.fullmatch(name.strip())
            if m:
                return int(m.group(1))
            if name.strip().lower() == "title":
                return 0
        ppr = style.element.find(q("w:pPr"))
        if ppr is not None:
            ol = ppr.find(q("w:outlineLvl"))
            if ol is not None and ol.get(q("w:val"), "").isdigit():
                v = int(ol.get(q("w:val")))
                # An explicit level stops inheritance; 9 means body text (e.g. "TOC Heading",
                # which is based on Heading 1 but must not appear in the TOC).
                return v + 1 if v < 9 else None
        style = style.base_style
        seen += 1
    return None


def headings(ctx: Context) -> list[tuple[int, str]]:
    """[(level, text)] for every heading paragraph in the body, in document order.

    A paragraph is a heading when its own outline level, or its style (or a base style),
    is "Heading N" / has an outline level. "Title" counts as level 0.
    """
    if "docx_headings" in ctx._cache:
        return ctx._cache["docx_headings"]
    doc = ctx.docx()
    styles = {s.style_id: s for s in doc.styles}
    default = doc.styles.default(WD_STYLE_TYPE.PARAGRAPH)
    level_cache: dict[str, int | None] = {}
    out = []
    for p in doc.element.body.iter(W_P):
        level = None
        ppr = p.find(q("w:pPr"))
        ps = ppr.find(q("w:pStyle")) if ppr is not None else None
        if ppr is not None:
            ol = ppr.find(q("w:outlineLvl"))
            if ol is not None and (ol.get(q("w:val")) or "").isdigit():
                v = int(ol.get(q("w:val")))
                level = v + 1 if v < 9 else None
        if level is None:
            sid = ps.get(q("w:val")) if ps is not None else None
            key = sid or "__default__"
            if key not in level_cache:
                st = styles.get(sid) if sid else default
                level_cache[key] = _style_level(st) if st is not None else None
            level = level_cache[key]
        if level is not None:
            text = norm_ws(w_text(p))
            if text:
                out.append((level, text))
    ctx._cache["docx_headings"] = out
    return out


@check("docx.has_heading", {**TEXT_MATCH_PROPS, "level": LEVEL}, **TEXT_OR_REGEX)
def has_heading(ctx: Context, p: dict[str, Any]) -> Outcome:
    m = TextMatcher(p)
    hs = headings(ctx)
    lvl = p.get("level")
    ok = any(m.matches(t) and (lvl is None or lv == lvl) for lv, t in hs)
    exp = m.describe() | ({"level": lvl} if lvl is not None else {})
    return Outcome(ok, exp, [{"level": lv, "text": t} for lv, t in first_n(hs, 50)])


@check("docx.heading_count", {**COUNT_PROPS, "level": LEVEL}, **COUNT_REQ)
def heading_count(ctx: Context, p: dict[str, Any]) -> Outcome:
    lvl = p.get("level")
    hs = [h for h in headings(ctx) if (h[0] == lvl if lvl is not None else h[0] >= 1)]
    exp = count_expected(p) | ({"level": lvl} if lvl is not None else {})
    return Outcome(count_ok(len(hs), p), exp, len(hs))


# ---- table of contents -------------------------------------------------------------------------

_PAGE_NUM_TAIL = re.compile(r"\s*\t\s*([0-9]+|[ivxlcdm]+)\s*$", re.IGNORECASE)


def _toc(ctx: Context) -> dict[str, Any]:
    """Find the first TOC field in the body and its pre-filled entries."""
    body = ctx.docx().element.body
    found: dict[str, Any] = {"present": False, "instr": None, "entries": []}
    stack: list[dict[str, Any]] = []  # open complex fields: {"instr", "sep"}

    def toc_active() -> bool:
        return any(f["sep"] and f["instr"].lstrip().upper().startswith("TOC") for f in stack)

    def in_pageref() -> bool:
        return any(f["instr"].lstrip().upper().startswith("PAGEREF") for f in stack)

    for para in body.iter(W_P):
        buf: list[str] = []
        touched = False
        for el in para.iter():
            t = el.tag
            if t == q("w:fldChar"):
                kind = el.get(q("w:fldCharType"))
                if kind == "begin":
                    stack.append({"instr": "", "sep": False})
                elif kind == "separate" and stack:
                    stack[-1]["sep"] = True
                    if stack[-1]["instr"].lstrip().upper().startswith("TOC") and not found["present"]:
                        found.update(present=True, instr=stack[-1]["instr"].strip())
                elif kind == "end" and stack:
                    closed = stack.pop()
                    if closed["instr"].lstrip().upper().startswith("TOC") and found["present"]:
                        if buf:
                            found["entries"].append("".join(buf))
                        return found
            elif t == q("w:instrText") and stack:
                stack[-1]["instr"] += el.text or ""
            elif t == q("w:fldSimple") and el.get(q("w:instr"), "").lstrip().upper().startswith("TOC"):
                if not found["present"]:
                    found.update(present=True, instr=el.get(q("w:instr")).strip())
            elif t == q("w:t") and toc_active() and not in_pageref():
                buf.append(el.text or "")
                touched = True
            elif t == q("w:tab") and el.getparent() is not None and el.getparent().tag == q("w:r") \
                    and toc_active() and not in_pageref():
                buf.append("\t")
        if touched and buf:
            found["entries"].append("".join(buf))
    if found["present"] and stack:
        found["unterminated"] = True
    return found


def _entry_title(raw: str) -> str:
    s = _PAGE_NUM_TAIL.sub("", raw)
    segs = [x for x in (norm_ws(seg) for seg in s.split("\t")) if x]
    return segs[-1] if segs else ""


def _toc_levels(instr: str) -> tuple[int, int]:
    m = re.search(r'\\o\s+"?\s*(\d)\s*-\s*(\d)', instr or "")
    return (int(m.group(1)), int(m.group(2))) if m else (1, 9)


@check("docx.has_toc", {"check_entries": BOOL, "ignore_case": BOOL,
                        "levels": {"type": "array", "items": LEVEL, "minItems": 2, "maxItems": 2}})
def has_toc(ctx: Context, p: dict[str, Any]) -> Outcome:
    toc = _toc(ctx)
    check_entries = p.get("check_entries", True)
    lo, hi = p["levels"] if "levels" in p else _toc_levels(toc["instr"])
    want = [t for lv, t in headings(ctx) if lo <= lv <= hi]
    got = [e for e in (_entry_title(x) for x in toc["entries"]) if e]
    exp: dict[str, Any] = {"toc_field": True}
    if check_entries:
        exp["entries"] = want
    act: dict[str, Any] = {"toc_field": toc["present"], "instr": toc["instr"], "entries": got}
    ok = toc["present"]
    if check_entries:
        if p.get("ignore_case"):
            ok = ok and [x.casefold() for x in got] == [x.casefold() for x in want]
        else:
            ok = ok and got == want
        ok = ok and bool(want)
    return Outcome(ok, exp, act)


# ---- tables ------------------------------------------------------------------------------------


def _tables(ctx: Context) -> list[dict[str, Any]]:
    out = []
    for tbl in ctx.docx().element.body.iter(q("w:tbl")):
        rows = tbl.findall(q("w:tr"))
        cols = 0
        for tr in rows:
            n = 0
            for tc in tr.findall(q("w:tc")):
                span = tc.find(f"{q('w:tcPr')}/{q('w:gridSpan')}")
                n += int(span.get(q("w:val"), "1")) if span is not None else 1
            cols = max(cols, n)
        header = [norm_ws(w_text(tc)) for tc in rows[0].findall(q("w:tc"))] if rows else []
        out.append({"rows": len(rows), "cols": cols, "header": header})
    return out


@check("docx.has_table", {"min_rows": POS_INT, "min_cols": POS_INT, "header_contains": STR_OR_LIST,
                          "ignore_case": BOOL, "min_count": POS_INT})
def has_table(ctx: Context, p: dict[str, Any]) -> Outcome:
    tables = _tables(ctx)
    hdr = p.get("header_contains")
    hdr = [hdr] if isinstance(hdr, str) else (hdr or [])
    fold = (lambda s: s.casefold()) if p.get("ignore_case") else (lambda s: s)

    def fits(t: dict[str, Any]) -> bool:
        cells = {fold(c) for c in t["header"]}
        return (t["rows"] >= p.get("min_rows", 1) and t["cols"] >= p.get("min_cols", 1)
                and all(fold(norm_ws(h)) in cells for h in hdr))

    n = sum(1 for t in tables if fits(t))
    exp = {k: p[k] for k in ("min_rows", "min_cols", "header_contains", "min_count") if k in p}
    return Outcome(n >= p.get("min_count", 1), exp or {"min_count": 1},
                   {"matching": n, "tables": first_n(tables, 20)})


# ---- text --------------------------------------------------------------------------------------

SCOPE = {"enum": ["body", "all"]}


def doc_text(ctx: Context, scope: str = "body") -> str:
    key = f"docx_text_{scope}"
    if key not in ctx._cache:
        texts = w_paragraph_texts(ctx.docx().element.body)
        if scope == "all":
            zf = ctx.zip()
            for part in parts_matching(zf, r"word/(header\d*|footer\d*|footnotes|endnotes|comments)\.xml"):
                texts.extend(w_paragraph_texts(read_xml(zf, part)))
        ctx._cache[key] = "\n".join(texts)
    return ctx._cache[key]


@check("docx.contains_text", {**TEXT_MATCH_PROPS, "scope": SCOPE}, **TEXT_OR_REGEX)
def contains_text(ctx: Context, p: dict[str, Any]) -> Outcome:
    m = TextMatcher(p, "contains")
    text = doc_text(ctx, p.get("scope", "body"))
    ok = m.matches(text)
    return Outcome(ok, m.describe() | {"present": True}, {"present": ok, "chars": len(text)})


@check("docx.not_contains_text", {**TEXT_MATCH_PROPS, "scope": SCOPE}, **TEXT_OR_REGEX)
def not_contains_text(ctx: Context, p: dict[str, Any]) -> Outcome:
    m = TextMatcher(p, "contains")
    text = doc_text(ctx, p.get("scope", "body"))
    found = m.matches(text)
    return Outcome(not found, m.describe() | {"present": False}, {"present": found})


# ---- styles, layout ----------------------------------------------------------------------------


@check("docx.style_used", {"style": NONEMPTY_STR}, required=("style",))
def style_used(ctx: Context, p: dict[str, Any]) -> Outcome:
    doc = ctx.docx()
    by_id = {s.style_id: s.name for s in doc.styles}
    default_para = doc.styles.default(WD_STYLE_TYPE.PARAGRAPH)
    roots = [doc.element.body]
    zf = ctx.zip()
    roots += [read_xml(zf, n) for n in parts_matching(zf, r"word/(header\d*|footer\d*|footnotes|endnotes)\.xml")]
    used: set[str] = set()
    for root in roots:
        for el in root.iter():
            if el.tag in (q("w:pStyle"), q("w:rStyle"), q("w:tblStyle")):
                sid = el.get(q("w:val"), "")
                used.add(by_id.get(sid, sid))
            elif el.tag == W_P and default_para is not None:
                ppr = el.find(q("w:pPr"))
                if (ppr is None or ppr.find(q("w:pStyle")) is None) and w_text(el).strip():
                    used.add(default_para.name)
    want = p["style"].strip().casefold()
    ids = {name: sid for sid, name in by_id.items()}
    ok = any(want in (u.casefold(), (ids.get(u) or "").casefold()) for u in used)
    return Outcome(ok, p["style"], sorted(used))


@check("docx.page_orientation", {"orientation": {"enum": ["portrait", "landscape"]}, "section": POS_INT},
       required=("orientation",))
def page_orientation(ctx: Context, p: dict[str, Any]) -> Outcome:
    secs = ctx.docx().sections
    orients = []
    for s in secs:
        w, h = s.page_width, s.page_height
        if w is not None and h is not None:
            orients.append("landscape" if w > h else "portrait")
        else:
            orients.append("landscape" if s.orientation == 1 else "portrait")
    if "section" in p:
        i = p["section"] - 1
        if i >= len(orients):
            raise CheckError(f"document has {len(orients)} section(s), not {p['section']}")
        return Outcome(orients[i] == p["orientation"], {"section": p["section"], "orientation": p["orientation"]},
                       orients[i])
    return Outcome(bool(orients) and all(o == p["orientation"] for o in orients),
                   {"all_sections": p["orientation"]}, orients)


@check("docx.header_contains", {**TEXT_MATCH_PROPS}, **TEXT_OR_REGEX)
def header_contains(ctx: Context, p: dict[str, Any]) -> Outcome:
    zf = ctx.zip()
    m = TextMatcher(p, "contains")
    texts = ["\n".join(w_paragraph_texts(read_xml(zf, n))) for n in parts_matching(zf, r"word/header\d*\.xml")]
    return Outcome(any(m.matches(t) for t in texts), m.describe(), [norm_ws(t) for t in texts])


_PAGE_FIELD = re.compile(r"\s*PAGE\b", re.IGNORECASE)


@check("docx.footer_has_page_number", {})
def footer_has_page_number(ctx: Context, p: dict[str, Any]) -> Outcome:
    zf = ctx.zip()
    footers = parts_matching(zf, r"word/footer\d*\.xml")
    instrs = sorted({i for n in footers for i in field_instructions(read_xml(zf, n))})
    ok = any(_PAGE_FIELD.match(i) for i in instrs)
    return Outcome(ok, "PAGE field in a footer", {"footers": len(footers), "fields": instrs})
