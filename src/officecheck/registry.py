"""Check registry: every check type registers its parameter schema and implementation."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Callable

FORMATS = ("docx", "xlsx", "pptx", "pdf", "any")


@dataclass
class Outcome:
    """What a check function returns."""

    passed: bool
    expected: Any
    actual: Any


class CheckError(Exception):
    """Raised inside a check when it cannot be evaluated (wrong file type, bad reference...).

    The runner turns it into a failed check whose ``actual`` is ``{"error": message}``.
    """


@dataclass(frozen=True)
class CheckDef:
    type: str
    func: Callable[..., Outcome]
    properties: dict[str, Any]
    required: tuple[str, ...] = ()
    extra: dict[str, Any] = field(default_factory=dict)  # extra schema keywords (anyOf, oneOf...)

    @property
    def fmt(self) -> str:
        return self.type.split(".", 1)[0]

    def schema(self) -> dict[str, Any]:
        props = {
            "type": {"const": self.type},
            "id": {"type": "string", "minLength": 1},
            "weight": {"type": "number", "exclusiveMinimum": 0},
            "description": {"type": "string"},
            **self.properties,
        }
        schema: dict[str, Any] = {
            "type": "object",
            "properties": props,
            "required": ["type", *self.required],
            "additionalProperties": False,
        }
        schema.update(self.extra)
        return schema


REGISTRY: dict[str, CheckDef] = {}


def check(
    type_: str,
    properties: dict[str, Any] | None = None,
    required: tuple[str, ...] | list[str] = (),
    **extra: Any,
) -> Callable[[Callable[..., Outcome]], Callable[..., Outcome]]:
    """Register ``func(ctx, params) -> Outcome`` as check type ``type_`` (e.g. ``"docx.has_heading"``)."""
    fmt = type_.split(".", 1)[0]
    if fmt not in FORMATS or "." not in type_:
        raise ValueError(f"check type must be '<format>.<name>' with format in {FORMATS}: {type_!r}")

    def deco(func: Callable[..., Outcome]) -> Callable[..., Outcome]:
        if type_ in REGISTRY:
            raise ValueError(f"duplicate check type {type_!r}")
        REGISTRY[type_] = CheckDef(type_, func, dict(properties or {}), tuple(required), extra)
        return func

    return deco


# ---- reusable schema fragments -------------------------------------------------------------

STR = {"type": "string"}
NONEMPTY_STR = {"type": "string", "minLength": 1}
BOOL = {"type": "boolean"}
POS_INT = {"type": "integer", "minimum": 1}
NONNEG_INT = {"type": "integer", "minimum": 0}
NUMBER = {"type": "number"}
STR_OR_LIST = {"oneOf": [NONEMPTY_STR, {"type": "array", "items": NONEMPTY_STR, "minItems": 1}]}

# A text matcher: exactly one of text / regex, plus ignore_case.
TEXT_MATCH_PROPS = {"text": NONEMPTY_STR, "regex": NONEMPTY_STR, "ignore_case": BOOL}
TEXT_OR_REGEX = {"oneOf": [{"required": ["text"]}, {"required": ["regex"]}]}

# Count bounds: any combination of equals / min / max, at least one.
COUNT_PROPS = {"equals": NONNEG_INT, "min": NONNEG_INT, "max": NONNEG_INT}
COUNT_REQ = {"anyOf": [{"required": ["equals"]}, {"required": ["min"]}, {"required": ["max"]}]}
