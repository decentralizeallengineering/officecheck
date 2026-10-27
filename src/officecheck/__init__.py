"""officecheck: deterministic grading of office files (docx, xlsx, pptx, pdf) against a JSON spec."""

from .core import grade
from .registry import REGISTRY, CheckError, Outcome, check
from .spec import SpecError, load_spec, result_schema, spec_schema, validate_spec

__all__ = [
    "REGISTRY",
    "CheckError",
    "Outcome",
    "SpecError",
    "check",
    "grade",
    "load_spec",
    "result_schema",
    "spec_schema",
    "validate_spec",
]
__version__ = "0.1.0"
