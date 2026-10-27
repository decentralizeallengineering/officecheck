import json
from pathlib import Path

from officecheck import spec_schema, load_spec

ROOT = Path(__file__).resolve().parent.parent


def test_committed_schema_is_current():
    committed = json.loads((ROOT / "schema" / "spec.schema.json").read_text(encoding="utf-8"))
    assert committed == spec_schema(), "regenerate with: officecheck --print-schema > schema/spec.schema.json"


def test_example_specs_are_valid():
    for p in (ROOT / "examples").glob("*.json"):
        load_spec(p)
