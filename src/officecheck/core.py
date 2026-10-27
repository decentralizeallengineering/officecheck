"""Run a validated spec against a file and build the result dict."""

from __future__ import annotations

import os
from typing import Any

from .context import Context
from .registry import REGISTRY, CheckError, Outcome
from .spec import load_spec
from .util import jsonable


def check_id(index: int, chk: dict[str, Any]) -> str:
    return chk.get("id") or f"{index + 1}:{chk['type']}"


def run_check(ctx: Context, chk: dict[str, Any]) -> Outcome:
    cdef = REGISTRY[chk["type"]]
    params = {k: v for k, v in chk.items() if k not in ("type", "id", "weight", "description")}
    try:
        return cdef.func(ctx, params)
    except CheckError as e:
        return Outcome(False, params or chk["type"], {"error": str(e)})
    except Exception as e:  # noqa: BLE001 - a crashing check is a failed check, never a crash
        return Outcome(False, params or chk["type"], {"error": f"{type(e).__name__}: {e}"})


def grade(path: str | os.PathLike[str], spec: str | os.PathLike[str] | dict[str, Any]) -> dict[str, Any]:
    """Grade ``path`` against ``spec`` (dict, JSON text or path to a JSON file).

    Returns ``{"pass": bool, "score": float, "checks": [{"id", "pass", "expected", "actual"}]}``.
    ``score`` is the weight-averaged fraction of passing checks (weights default to 1);
    ``pass`` is ``score >= threshold`` (threshold defaults to 1.0, i.e. every check passes).
    Raises ``SpecError`` if the spec is invalid. Never raises because of the file itself.
    """
    from . import checks  # noqa: F401  - populate the registry

    spec = load_spec(spec)
    results = []
    total_w = passed_w = 0.0
    with Context(path) as ctx:
        for i, chk in enumerate(spec["checks"]):
            out = run_check(ctx, chk)
            w = float(chk.get("weight", 1.0))
            total_w += w
            if out.passed:
                passed_w += w
            results.append(
                {
                    "id": check_id(i, chk),
                    "pass": bool(out.passed),
                    "expected": jsonable(out.expected),
                    "actual": jsonable(out.actual),
                }
            )
    score = round(passed_w / total_w, 6) if total_w else 0.0
    threshold = float(spec.get("threshold", 1.0))
    ok = all(r["pass"] for r in results) if threshold >= 1.0 else score >= threshold
    return {"pass": ok, "score": score, "checks": results}
