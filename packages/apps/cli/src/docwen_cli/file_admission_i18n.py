"""CLI presentation adapter for stable file-admission diagnostics.

Core owns file inspection, stable diagnostic codes, and admission decisions.
This module owns only the CLI projection of those frozen facts into the active
locale; it must never infer or change an admission decision.
"""

from __future__ import annotations

from typing import Any, Protocol

from docwen_cli.i18n import cli_t
from docwen_core.models.file_inspection import file_admission_message_key


class FileAdmissionDiagnostic(Protocol):
    """Presentation-facing subset of Core's immutable inspection fact."""

    @property
    def declared_format(self) -> str: ...

    @property
    def detected_format(self) -> str: ...

    @property
    def warning_code(self) -> str: ...

    @property
    def warning_message(self) -> str: ...

    @property
    def reason_code(self) -> str: ...

    @property
    def reason_message(self) -> str: ...

    @property
    def warnings(self) -> tuple[dict[str, Any], ...]: ...


def render_file_admission_code(
    code: str,
    *,
    declared_format: str,
    detected_format: str,
    fallback: str = "",
) -> str:
    """Render one known code without exposing Core's English fallback text."""

    normalized_code = str(code or "").strip().upper()
    key = file_admission_message_key(normalized_code)
    if key is None:
        return fallback

    declared = str(declared_format or "unknown").upper()
    detected = str(detected_format or "unknown").upper()
    translated = cli_t(
        key,
        default="",
        declared_format=declared,
        detected_format=detected,
    )
    if translated and translated != key:
        return translated

    # A packaged CLI normally has the locale catalogue.  If that catalogue is
    # unavailable, retain the stable code and facts instead of leaking Core's
    # implementation-language message as if it were localized presentation.
    return f"[{normalized_code}] {declared} -> {detected}"


def render_file_inspection_message(
    inspection: FileAdmissionDiagnostic,
    *,
    prefer_reason: bool = False,
) -> str:
    """Render the most specific localized diagnostic for one inspection."""

    warning_code = "" if prefer_reason else str(getattr(inspection, "warning_code", ""))
    reason_code = str(getattr(inspection, "reason_code", ""))
    fallback = (
        str(getattr(inspection, "reason_message", ""))
        if prefer_reason
        else str(getattr(inspection, "warning_message", "") or getattr(inspection, "reason_message", ""))
    )
    # Confirmation uses one generic reason code plus a more specific warning
    # code.  Prefer the specific localized diagnostic when the generic reason
    # has no catalogue entry.
    for code in (reason_code, warning_code) if prefer_reason else (warning_code, reason_code):
        if file_admission_message_key(code) is not None:
            return render_file_admission_code(
                code,
                declared_format=str(getattr(inspection, "declared_format", "unknown")),
                detected_format=str(getattr(inspection, "detected_format", "unknown")),
                fallback=fallback,
            )
    for code in (reason_code, warning_code) if prefer_reason else (warning_code, reason_code):
        normalized_code = str(code or "").strip().upper()
        if normalized_code:
            declared = str(getattr(inspection, "declared_format", "unknown") or "unknown").upper()
            detected = str(getattr(inspection, "detected_format", "unknown") or "unknown").upper()
            return f"[{normalized_code}] {declared} -> {detected}"
    return fallback


def render_file_inspection_warning(inspection: FileAdmissionDiagnostic) -> str:
    """Render the localized admission warning plus non-admission diagnostics."""

    primary_code = str(getattr(inspection, "warning_code", "")).strip()
    parts = [render_file_inspection_message(inspection)]
    for item in getattr(inspection, "warnings", ()):
        code = str(item.get("code", "")).strip()
        if code and code == primary_code:
            continue
        message = str(item.get("message", "")).strip()
        if not message:
            continue
        parts.append(f"[{code}] {message}" if code else message)
    return " ".join(part for part in parts if part)


def detected_format_acceptance_hint() -> str:
    """Return a localized explanation for the explicit CLI acceptance flag."""

    key = "main_window.file_admission_confirm_action"
    action = cli_t(key, default="")
    return f"--use-detected-format: {action}" if action and action != key else "--use-detected-format"


__all__ = [
    "detected_format_acceptance_hint",
    "render_file_admission_code",
    "render_file_inspection_message",
    "render_file_inspection_warning",
]
