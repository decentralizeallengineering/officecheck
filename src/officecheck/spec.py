"""Spec loading and JSON-schema validation."""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import jsonschema

from .registry import REGISTRY


class SpecError(ValueError):
    """The spec is not valid JSON or does not match the spec schema."""


def spec_schema() -> dict[str, Any]:
    """The JSON Schema (draft 2020-12) every spec must satisfy."""
    from . import checks  # noqa: F401  - populate the registry

    types = sorted(REGISTRY)
    return {
        "$schema": "https://json-schema.org/draft/2020-12/schema",
        "$id": "https://officecheck.invalid/spec.schema.json",
        "title": "officecheck spec",
        "type": "object",
        "properties": {
            "version": {"const": 1},
            "description": {"type": "string"},
            "threshold": {"type": "number", "minimum": 0, "maximum": 1},
            "checks": {
                "type": "array",
                "minItems": 1,
                "items": {
                    "type": "object",
                    "properties": {"type": {"enum": types}},
                    "required": ["type"],
                    "allOf": [
                        {
                            "if": {"properties": {"type": {"const": t}}, "required": ["type"]},
                            "then": REGISTRY[t].schema(),
                        }
                        for t in types
                    ],
                },
            },
        },
        "required": ["checks"],
        "additionalProperties": False,
    }


def result_schema() -> dict[str, Any]:
    """JSON Schema of the grading result."""
    return {
        "$schema": "https://json-schema.org/draft/2020-12/schema",
        "type": "object",
        "properties": {
            "pass": {"type": "boolean"},
            "score": {"type": "number", "minimum": 0, "maximum": 1},
            "checks": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "id": {"type": "string"},
                        "pass": {"type": "boolean"},
                        "expected": {},
                        "actual": {},
                    },
                    "required": ["id", "pass", "expected", "actual"],
                    "additionalProperties": False,
                },
            },
        },
        "required": ["pass", "score", "checks"],
        "additionalProperties": False,
    }


_VALIDATOR: jsonschema.protocols.Validator | None = None


def _validator() -> jsonschema.protocols.Validator:
    global _VALIDATOR
    if _VALIDATOR is None:
        schema = spec_schema()
        cls = jsonschema.validators.validator_for(schema)
        cls.check_schema(schema)
        _VALIDATOR = cls(schema)
    return _VALIDATOR


def _fmt_error(err: jsonschema.ValidationError) -> str:
    where = "/".join(str(p) for p in err.absolute_path) or "(root)"
    msg = err.message
    subs = err.validator_value if err.validator in ("oneOf", "anyOf") else None
    if isinstance(subs, list) and subs and all(set(s) == {"required"} for s in subs):
        options = " | ".join("+".join(s["required"]) for s in subs)
        how = "exactly one" if err.validator == "oneOf" else "at least one"
        msg = f"needs {how} of: {options}"
    elif err.validator == "enum" and err.absolute_path and err.absolute_path[-1] == "type":
        msg = f"unknown check type {err.instance!r}"
    return f"{where}: {msg}"


def validate_spec(spec: Any) -> dict[str, Any]:
    """Validate a parsed spec; raise SpecError listing every problem. Returns the spec."""
    errors = sorted(_validator().iter_errors(spec), key=lambda e: list(map(str, e.absolute_path)))
    if errors:
        # The best-matching leaf errors are the readable ones; if/then wrappers add noise.
        leaves = []
        for e in errors:
            leaves.extend(_leaf_errors(e))
        msgs = list(dict.fromkeys(_fmt_error(e) for e in leaves))
        raise SpecError("invalid spec:\n  " + "\n  ".join(msgs))
    ids = [c.get("id") for c in spec["checks"] if "id" in c]
    dup = sorted({i for i in ids if ids.count(i) > 1})
    if dup:
        raise SpecError(f"invalid spec:\n  checks: duplicate id(s): {', '.join(dup)}")
    return spec


def _leaf_errors(err: jsonschema.ValidationError) -> list[jsonschema.ValidationError]:
    if err.context and err.validator in ("allOf",):
        out = []
        for sub in err.context:
            out.extend(_leaf_errors(sub))
        return out
    return [err]


def load_spec(source: str | os.PathLike[str] | dict[str, Any]) -> dict[str, Any]:
    """Load (from a path, a JSON string or a dict) and validate a spec."""
    if isinstance(source, dict):
        spec = source
    else:
        text = str(source)
        if isinstance(source, os.PathLike) or not text.lstrip().startswith("{"):
            try:
                text = Path(source).read_text(encoding="utf-8-sig")
            except OSError as e:
                raise SpecError(f"cannot read spec: {e}") from e
        try:
            spec = json.loads(text)
        except json.JSONDecodeError as e:
            raise SpecError(f"spec is not valid JSON: {e}") from e
    return validate_spec(spec)
