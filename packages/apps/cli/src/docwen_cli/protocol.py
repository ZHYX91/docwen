"""Protocol 3 primitives shared by every machine-facing CLI command."""

from __future__ import annotations

from enum import StrEnum
from typing import Any

from docwen_cli.error_registry import ERROR_SPECS, public_error_code
from docwen_core.version import PRODUCT_VERSION

PROTOCOL_VERSION = 3


class ErrorCategory(StrEnum):
    """Stable protocol categories; detailed error codes may evolve additively."""

    INTERNAL = "internal"
    INVALID_INPUT = "invalid_input"
    NOT_FOUND = "not_found"
    DEPENDENCY = "dependency"
    SECURITY = "security"
    UNAVAILABLE = "unavailable"
    CONFLICT = "conflict"
    TIMEOUT = "timeout"
    PARTIAL_FAILURE = "partial_failure"
    CANCELLED = "cancelled"


def category_for_error_code(code: str) -> ErrorCategory:
    """Return the stable category for a detailed protocol error code."""

    spec = ERROR_SPECS.get(public_error_code(code))
    return ErrorCategory(spec.category) if spec is not None else ErrorCategory.INTERNAL


def make_envelope(
    *,
    command: str,
    success: bool,
    data: Any | None = None,
    error: dict[str, Any] | None = None,
    warnings: list[Any] | None = None,
    meta: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Build the exact top-level protocol 3 envelope."""

    if success and error is not None:
        raise ValueError("A successful protocol 3 envelope must use error=null")
    if not success and not isinstance(error, dict):
        raise ValueError("A failed protocol 3 envelope must include a typed error object")

    return {
        "protocol_version": PROTOCOL_VERSION,
        "product_version": PRODUCT_VERSION,
        "success": success,
        "command": command,
        "data": data if data is not None else {},
        "error": error,
        "warnings": list(warnings or []),
        "meta": dict(meta or {}),
    }


__all__ = [
    "PROTOCOL_VERSION",
    "ErrorCategory",
    "category_for_error_code",
    "make_envelope",
]
