"""Command line: ``officecheck <file> <spec.json>``.

Exit status: 0 = pass, 1 = fail, 2 = usage error or invalid spec.
"""

from __future__ import annotations

import argparse
import json
import shutil
import sys
from typing import Any, Sequence

from . import __version__
from .core import grade
from .spec import SpecError, spec_schema


def _cell(v: Any) -> str:
    if isinstance(v, str):
        return v
    return json.dumps(v, ensure_ascii=False, sort_keys=False, default=str)


def _clip(s: str, width: int) -> str:
    s = " ".join(s.split())
    return s if len(s) <= width else s[: max(0, width - 3)] + "..."


def render_table(result: dict[str, Any], width: int | None = None) -> str:
    width = width or shutil.get_terminal_size((120, 24)).columns
    rows = [(r["id"], "PASS" if r["pass"] else "FAIL", _cell(r["expected"]), _cell(r["actual"]))
            for r in result["checks"]]
    id_w = min(max([len("CHECK")] + [len(r[0]) for r in rows]), 40)
    rest = max(width - id_w - 4 - 3 * 3, 30)
    exp_w = rest // 2
    act_w = rest - exp_w
    head = ("CHECK", "", "EXPECTED", "ACTUAL")
    lines = []
    for i, (cid, st, exp, act) in enumerate([head, *rows]):
        lines.append(f"{_clip(cid, id_w):<{id_w}} | {st:<4} | {_clip(exp, exp_w):<{exp_w}} | {_clip(act, act_w)}".rstrip())
        if i == 0:
            lines.append("-" * min(width, id_w + 4 + exp_w + act_w + 9))
    n_pass = sum(r["pass"] for r in result["checks"])
    verdict = "PASS" if result["pass"] else "FAIL"
    lines.append("")
    lines.append(f"{verdict}  score={result['score']:.3f}  ({n_pass}/{len(result['checks'])} checks passed)")
    return "\n".join(lines)


def main(argv: Sequence[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        prog="officecheck",
        description="Grade an office file (docx/xlsx/pptx/pdf) against a JSON spec. "
                    "Exit 0 = pass, 1 = fail, 2 = usage error / invalid spec.",
    )
    ap.add_argument("file", nargs="?", help="file to grade")
    ap.add_argument("spec", nargs="?", help="spec JSON file")
    ap.add_argument("--json", action="store_true", help="print the result as JSON instead of a table")
    ap.add_argument("--print-schema", action="store_true", help="print the spec JSON Schema and exit")
    ap.add_argument("--version", action="version", version=f"officecheck {__version__}")
    args = ap.parse_args(argv)
    for stream in (sys.stdout, sys.stderr):  # never crash on a legacy Windows console code page
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(errors="replace")

    if args.print_schema:
        print(json.dumps(spec_schema(), indent=2))
        return 0
    if not args.file or not args.spec:
        ap.print_usage(sys.stderr)
        print("officecheck: error: FILE and SPEC are required", file=sys.stderr)
        return 2
    try:
        result = grade(args.file, args.spec)
    except SpecError as e:
        print(f"officecheck: {e}", file=sys.stderr)
        return 2
    if args.json:
        print(json.dumps(result, indent=2, ensure_ascii=False))
    else:
        print(render_table(result))
    return 0 if result["pass"] else 1


if __name__ == "__main__":
    sys.exit(main())
