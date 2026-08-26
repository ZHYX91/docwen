"""Stable process exit codes for CLI protocol 3."""

from __future__ import annotations

from enum import IntEnum

from docwen_cli.error_registry import ERROR_SPECS, public_error_code


class ExitCode(IntEnum):
    """Exit codes aligned with the protocol 3 error categories."""

    OK = 0
    INTERNAL_ERROR = 1
    INVALID_INPUT = 2
    NOT_FOUND = 3
    DEPENDENCY_MISSING = 4
    SECURITY_CHECK_FAILED = 5
    UNAVAILABLE = 6
    CONFLICT = 7
    TIMEOUT = 8
    PARTIAL_FAILURE = 9
    CANCELLED = 130


# ── error_code string → ExitCode mapping ────────────────────────────

ERROR_CODE_TO_EXIT_CODE: dict[str, ExitCode] = {code: ExitCode(spec.exit_code) for code, spec in ERROR_SPECS.items()}
ERROR_CODE_TO_EXIT_CODE["skipped_same_format"] = ExitCode.OK


def exit_code_from_error_code(code: str) -> ExitCode:
    """Map an error_code string to an ExitCode.

    Returns ``INTERNAL_ERROR`` for unrecognised codes.
    """
    return ERROR_CODE_TO_EXIT_CODE.get(public_error_code(code), ExitCode.INTERNAL_ERROR)
