"""Small shared helpers: text matching, count bounds, JSON-safe values."""

from __future__ import annotations

import datetime as _dt
import decimal
import math
import re
from typing import Any, Iterable

_WS = re.compile(r"\s+")


def norm_ws(s: str) -> str:
    """Collapse all whitespace runs (incl. NBSP, tabs, newlines) to one space and strip."""
    return _WS.sub(" ", s.replace(" ", " ")).strip()


class TextMatcher:
    """Matches strings against ``text`` (whitespace-normalised) or ``regex`` (``re.search``).

    mode="equals": the whole (normalised) string must equal ``text``.
    mode="contains": ``text`` must occur as a substring.
    A regex is always applied with ``re.search``; anchor it yourself for a full match.
    """

    def __init__(self, params: dict[str, Any], mode: str = "equals"):
        self.ignore_case = bool(params.get("ignore_case", False))
        self.mode = mode
        self.text = params.get("text")
        self.regex = params.get("regex")
        if self.regex is not None:
            self._rx = re.compile(self.regex, re.IGNORECASE if self.ignore_case else 0)
        elif self.text is None:
            raise ValueError("matcher needs text or regex")

    def matches(self, s: str | None) -> bool:
        if s is None:
            return False
        if self.regex is not None:
            return self._rx.search(s) is not None
        hay, needle = norm_ws(s), norm_ws(self.text)
        if self.ignore_case:
            hay, needle = hay.casefold(), needle.casefold()
        return hay == needle if self.mode == "equals" else needle in hay

    def describe(self) -> dict[str, Any]:
        d: dict[str, Any] = {"regex": self.regex} if self.regex is not None else {"text": self.text}
        if self.ignore_case:
            d["ignore_case"] = True
        return d


def count_ok(n: int, params: dict[str, Any]) -> bool:
    if "equals" in params and n != params["equals"]:
        return False
    if "min" in params and n < params["min"]:
        return False
    if "max" in params and n > params["max"]:
        return False
    return True


def count_expected(params: dict[str, Any]) -> dict[str, int]:
    return {k: params[k] for k in ("equals", "min", "max") if k in params}


def jsonable(v: Any) -> Any:
    """Convert values to something json.dumps accepts, deterministically."""
    if v is None or isinstance(v, (bool, int, str)):
        return v
    if isinstance(v, float):
        return v if math.isfinite(v) else str(v)
    if isinstance(v, decimal.Decimal):
        return float(v)
    if isinstance(v, (_dt.datetime, _dt.date, _dt.time)):
        return v.isoformat()
    if isinstance(v, _dt.timedelta):
        return v.total_seconds()
    if isinstance(v, dict):
        return {str(k): jsonable(x) for k, x in v.items()}
    if isinstance(v, (list, tuple, set, frozenset)):
        items = [jsonable(x) for x in v]
        if isinstance(v, (set, frozenset)):
            items.sort(key=repr)
        return items
    return str(v)


def first_n(items: Iterable[Any], n: int = 10) -> list[Any]:
    out = []
    for i, x in enumerate(items):
        if i >= n:
            break
        out.append(x)
    return out
