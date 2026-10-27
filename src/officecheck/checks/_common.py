"""Helpers shared by several check modules."""

from __future__ import annotations

from typing import Any

from ..ooxml import chart_type_matches
from ..registry import BOOL, NONEMPTY_STR, NONNEG_INT, POS_INT, Outcome
from ..util import TextMatcher, first_n

CHART_TYPES = ["bar", "column", "line", "pie", "doughnut", "area", "scatter", "bubble", "radar", "stock", "surface"]
CHART_PROPS = {"chart_type": {"enum": CHART_TYPES}, "series": NONNEG_INT, "min_series": NONNEG_INT,
               "title": NONEMPTY_STR, "title_regex": NONEMPTY_STR, "ignore_case": BOOL, "min_count": POS_INT}


def match_charts(charts: list[dict[str, Any]], p: dict[str, Any]) -> Outcome:
    tm = None
    if "title" in p or "title_regex" in p:
        tm = TextMatcher({"text": p.get("title"), "regex": p.get("title_regex"),
                          "ignore_case": p.get("ignore_case", False)})

    def fits(c: dict[str, Any]) -> bool:
        if "chart_type" in p and not chart_type_matches(p["chart_type"], c["types"]):
            return False
        if "series" in p and c["series"] != p["series"]:
            return False
        if "min_series" in p and c["series"] < p["min_series"]:
            return False
        return tm is None or tm.matches(c["title"])

    n = sum(1 for c in charts if fits(c))
    keys = ("sheet", "slide", "chart_type", "series", "min_series", "title", "title_regex", "min_count")
    exp = {k: p[k] for k in keys if k in p} or {"min_count": 1}
    return Outcome(n >= p.get("min_count", 1), exp, {"matching": n, "charts": first_n(charts, 20)})
