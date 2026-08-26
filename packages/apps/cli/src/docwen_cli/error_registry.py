"""Single source of truth for stable machine error semantics."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class ErrorSpec:
    category: str
    exit_code: int


def _spec(category: str, exit_code: int) -> ErrorSpec:
    return ErrorSpec(category=category, exit_code=exit_code)


ERROR_SPECS: dict[str, ErrorSpec] = {
    "invalid_arguments": _spec("invalid_input", 2),
    "invalid_input": _spec("invalid_input", 2),
    "invalid_path": _spec("invalid_input", 2),
    "invalid_control_action": _spec("invalid_input", 2),
    "control_message_too_large": _spec("invalid_input", 2),
    "file_not_found": _spec("invalid_input", 2),
    "file_empty": _spec("invalid_input", 2),
    "file_container_invalid": _spec("invalid_input", 2),
    "file_container_unsupported": _spec("invalid_input", 2),
    "file_container_unrecognized": _spec("invalid_input", 2),
    "file_content_unrecognized": _spec("invalid_input", 2),
    "file_format_confirmation_required": _spec("invalid_input", 2),
    "unsupported_format": _spec("invalid_input", 2),
    "resource_not_found": _spec("not_found", 3),
    "dependency_missing": _spec("dependency", 4),
    "security_check_failed": _spec("security", 5),
    "network_access_blocked": _spec("security", 5),
    "capability_unavailable": _spec("unavailable", 6),
    "gui_not_running": _spec("unavailable", 6),
    "gui_start_failed": _spec("unavailable", 6),
    "settings_section_unavailable": _spec("unavailable", 6),
    "unsupported_control_protocol": _spec("unavailable", 6),
    "invalid_control_response": _spec("unavailable", 6),
    "invalid_control_message": _spec("unavailable", 6),
    "gui_control_start_failed": _spec("unavailable", 6),
    "gui_control_endpoint_unavailable": _spec("unavailable", 6),
    "protocol_incompatible": _spec("unavailable", 6),
    "unsupported_route": _spec("unavailable", 6),
    "unsupported_numbering": _spec("unavailable", 6),
    "output_exists": _spec("conflict", 7),
    "output_collision": _spec("conflict", 7),
    "operation_timeout": _spec("timeout", 8),
    "control_timeout": _spec("timeout", 8),
    "batch_partial_failure": _spec("partial_failure", 9),
    "operation_cancelled": _spec("cancelled", 130),
    "conversion_failed": _spec("internal", 1),
    "gui_command_failed": _spec("internal", 1),
    "internal_error": _spec("internal", 1),
    "unknown_error": _spec("internal", 1),
}

_INTERNAL_TO_PUBLIC_ERROR_CODE = {
    "cancelled": "operation_cancelled",
}


def public_error_code(code: str) -> str:
    """Normalize internal result labels to canonical protocol 3 error codes."""
    return _INTERNAL_TO_PUBLIC_ERROR_CODE.get(code, code)


__all__ = ["ERROR_SPECS", "ErrorSpec", "public_error_code"]
