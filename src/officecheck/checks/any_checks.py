"""Checks that apply to any file type."""

from __future__ import annotations

from typing import Any

from ..context import Context
from ..masking import blobs
from ..registry import BOOL, NONEMPTY_STR, CheckError, Outcome, check
from ..util import norm_ws


@check("any.file_opens", {"format": {"enum": ["docx", "xlsx", "pptx", "pdf"]}})
def file_opens(ctx: Context, p: dict[str, Any]) -> Outcome:
    exp = {"opens": True} | ({"format": p["format"]} if "format" in p else {})
    try:
        ctx.open_native()
    except CheckError as e:
        return Outcome(False, exp, {"opens": False, "format": ctx.format, "error": str(e)})
    ok = "format" not in p or ctx.format == p["format"]
    return Outcome(ok, exp, {"opens": True, "format": ctx.format})


@check("any.size_under", {"max_bytes": {"type": "integer", "minimum": 1}}, required=("max_bytes",))
def size_under(ctx: Context, p: dict[str, Any]) -> Outcome:
    n = len(ctx.require_bytes())
    return Outcome(n < p["max_bytes"], {"max_bytes": p["max_bytes"]}, n)


def _norm(s: str, ignore_case: bool) -> str:
    s = norm_ws(s)
    return s.casefold() if ignore_case else s


@check("any.masked", {"strings": {"type": "array", "items": NONEMPTY_STR, "minItems": 1},
                      "ignore_case": BOOL}, required=("strings",))
def masked(ctx: Context, p: dict[str, Any]) -> Outcome:
    ic = p.get("ignore_case", True)
    needles = {s: _norm(s, ic) for s in p["strings"]}
    found: dict[str, list[str]] = {}
    for loc, text in blobs(ctx.require_bytes()):
        hay = _norm(text, ic)
        for s, n in needles.items():
            if n and n in hay:
                locs = found.setdefault(s, [])
                if loc not in locs and len(locs) < 10:
                    locs.append(loc)
    exp = {"absent": p["strings"], "ignore_case": ic}
    return Outcome(not found, exp, {"found": [{"string": s, "locations": found[s]} for s in p["strings"] if s in found]})
