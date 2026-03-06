from __future__ import annotations

from dataclasses import dataclass

from docwen.services.error_codes import (
    ERROR_CODE_CONVERSION_FAILED,
    ERROR_CODE_DEPENDENCY_MISSING,
    ERROR_CODE_INVALID_INPUT,
    ERROR_CODE_NOT_IMPLEMENTED,
    ERROR_CODE_OPERATION_CANCELLED,
    ERROR_CODE_SECURITY_CHECK_FAILED,
    ERROR_CODE_SKIPPED_SAME_FORMAT,
    ERROR_CODE_STRATEGY_NOT_FOUND,
    ERROR_CODE_UNKNOWN_ERROR,
    ERROR_CODE_UNSUPPORTED_FORMAT,
)


@dataclass(frozen=True, slots=True)
class ErrorDefinition:
    code: str
    exit_code: int
    retryable: bool
    kind: str


ERROR_REGISTRY: dict[str, ErrorDefinition] = {
    ERROR_CODE_UNKNOWN_ERROR: ErrorDefinition(
        code=ERROR_CODE_UNKNOWN_ERROR, exit_code=1, retryable=True, kind="internal"
    ),
    ERROR_CODE_INVALID_INPUT: ErrorDefinition(
        code=ERROR_CODE_INVALID_INPUT, exit_code=2, retryable=False, kind="input"
    ),
    ERROR_CODE_STRATEGY_NOT_FOUND: ErrorDefinition(
        code=ERROR_CODE_STRATEGY_NOT_FOUND, exit_code=3, retryable=False, kind="strategy"
    ),
    ERROR_CODE_DEPENDENCY_MISSING: ErrorDefinition(
        code=ERROR_CODE_DEPENDENCY_MISSING, exit_code=4, retryable=True, kind="dependency"
    ),
    ERROR_CODE_SECURITY_CHECK_FAILED: ErrorDefinition(
        code=ERROR_CODE_SECURITY_CHECK_FAILED, exit_code=5, retryable=False, kind="security"
    ),
    ERROR_CODE_UNSUPPORTED_FORMAT: ErrorDefinition(
        code=ERROR_CODE_UNSUPPORTED_FORMAT, exit_code=2, retryable=False, kind="input"
    ),
    ERROR_CODE_NOT_IMPLEMENTED: ErrorDefinition(
        code=ERROR_CODE_NOT_IMPLEMENTED, exit_code=1, retryable=False, kind="internal"
    ),
    ERROR_CODE_CONVERSION_FAILED: ErrorDefinition(
        code=ERROR_CODE_CONVERSION_FAILED, exit_code=1, retryable=True, kind="conversion"
    ),
    ERROR_CODE_OPERATION_CANCELLED: ErrorDefinition(
        code=ERROR_CODE_OPERATION_CANCELLED, exit_code=1, retryable=True, kind="cancelled"
    ),
    ERROR_CODE_SKIPPED_SAME_FORMAT: ErrorDefinition(
        code=ERROR_CODE_SKIPPED_SAME_FORMAT, exit_code=0, retryable=False, kind="skipped"
    ),
}


def get_error_definition(error_code: str | None) -> ErrorDefinition:
    if not error_code:
        return ERROR_REGISTRY[ERROR_CODE_UNKNOWN_ERROR]
    return ERROR_REGISTRY.get(error_code) or ERROR_REGISTRY[ERROR_CODE_UNKNOWN_ERROR]
