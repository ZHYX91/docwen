"""Localized rendering for stable file-admission warning and reason codes."""

from __future__ import annotations

from collections.abc import Mapping
from typing import Any, Protocol

from docwen_core.models.file_inspection import file_admission_message_key
from docwen_gui.i18n import t


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


_FORMAT_NOTICE_CODES = frozenset(
    {
        "FILE_FORMAT_COMPATIBLE_TEXT",
        "FILE_FORMAT_SAME_FAMILY_MISMATCH",
        "FILE_FORMAT_CROSS_FAMILY_MISMATCH",
        "FILE_EXTENSION_UNSUPPORTED",
    }
)


def render_file_format_notice(file_ref: Any) -> str:
    """Return a compact actual-format notice for one admitted mismatch.

    This is presentation-only.  Admission and confirmation semantics remain
    owned by the frozen inspection fact.
    """

    try:
        from docwen_core.models import FILE_INSPECTION_METADATA_KEY
    except Exception:
        return ""

    metadata = getattr(file_ref, "metadata", {})
    inspection = metadata.get(FILE_INSPECTION_METADATA_KEY) if isinstance(metadata, dict) else None
    if not isinstance(inspection, Mapping):
        return ""

    code = str(inspection.get("warning_code", "") or "").strip().upper()
    if code not in _FORMAT_NOTICE_CODES:
        return ""

    detected = str(inspection.get("detected_format", "") or getattr(file_ref, "format", "") or "").strip()
    if not detected or detected.casefold() == "unknown":
        return ""
    return t(
        "file_admission.actual_format",
        "Actual format: {detected_format}",
        detected_format=detected.upper(),
    )


_ENGLISH_FALLBACKS: dict[str, str] = {
    "FILE_FORMAT_COMPATIBLE_TEXT": (
        "The filename declares {declared_format}, while the content was detected as "
        "{detected_format}; both use the same text workflow."
    ),
    "FILE_FORMAT_SAME_FAMILY_MISMATCH": (
        "The filename declares {declared_format}, while the content was detected as "
        "{detected_format}. Both formats use the same processing family, so the detected format will be used."
    ),
    "FILE_FORMAT_CROSS_FAMILY_MISMATCH": (
        "The filename declares {declared_format}, while the content was detected as "
        "{detected_format}. The formats use different processing families; confirm processing as the detected format."
    ),
    "FILE_EXTENSION_UNSUPPORTED": (
        "The filename extension declares unsupported format {declared_format}, while the content was detected as "
        "{detected_format}. Confirm processing as the detected format."
    ),
    "FILE_EMPTY": "The selected file is empty.",
    "FILE_CONTAINER_INVALID": (
        "The file container is corrupt or structurally invalid (declared {declared_format}, "
        "detected {detected_format})."
    ),
    "FILE_CONTAINER_UNSUPPORTED": (
        "The detected {detected_format} container is not a supported document package (declared {declared_format})."
    ),
    "FILE_CONTAINER_UNRECOGNIZED": (
        "The file container could not be recognized (declared {declared_format}, detected {detected_format})."
    ),
    "FILE_CONTENT_UNRECOGNIZED": (
        "The file content could not be recognized (declared {declared_format}, detected {detected_format})."
    ),
    "FILE_READ_ERROR": ("The file could not be read (declared {declared_format}, detected {detected_format})."),
    "UNSUPPORTED_FORMAT": (
        "The detected content format {detected_format} is not supported (declared {declared_format})."
    ),
}


def render_file_admission_code(
    code: str,
    *,
    declared_format: str,
    detected_format: str,
    fallback: str = "",
) -> str:
    """Render one stable code without exposing Core's English message text."""

    normalized_code = str(code or "").strip().upper()
    key = file_admission_message_key(normalized_code)
    if key is None:
        return fallback
    default = _ENGLISH_FALLBACKS[normalized_code]
    return t(
        key,
        default,
        declared_format=str(declared_format or "unknown").upper(),
        detected_format=str(detected_format or "unknown").upper(),
    )


def render_file_admission_message(
    *,
    warning_code: str = "",
    reason_code: str = "",
    declared_format: str,
    detected_format: str,
    fallback: str = "",
) -> str:
    """Render the most specific localized admission diagnostic available."""

    for code in (warning_code, reason_code):
        if file_admission_message_key(code) is not None:
            return render_file_admission_code(
                code,
                declared_format=declared_format,
                detected_format=detected_format,
                fallback=fallback,
            )
    return fallback


def render_file_inspection_message(
    inspection: FileAdmissionDiagnostic,
    *,
    prefer_reason: bool = False,
) -> str:
    """Render one inspection diagnostic at the GUI presentation boundary.

    Core keeps stable codes plus English fallbacks for non-GUI consumers.  The
    GUI never re-decides admission here; it only projects the already-frozen
    fact into the active locale.
    """

    reason_code = str(getattr(inspection, "reason_code", ""))
    reason_message = str(getattr(inspection, "reason_message", ""))
    declared_format = str(getattr(inspection, "declared_format", "unknown"))
    detected_format = str(getattr(inspection, "detected_format", "unknown"))
    if prefer_reason:
        # Cross-family and unknown-extension confirmation share one generic
        # reason code.  The warning code carries the specific localized facts.
        warning_code = str(getattr(inspection, "warning_code", ""))
        return render_file_admission_message(
            warning_code=warning_code,
            reason_code=reason_code,
            declared_format=declared_format,
            detected_format=detected_format,
            fallback=reason_message,
        )

    rendered_warnings: list[str] = []
    raw_warnings = getattr(inspection, "warnings", ())
    if isinstance(raw_warnings, (tuple, list)):
        for raw_warning in raw_warnings:
            if not isinstance(raw_warning, Mapping):
                continue
            code = str(raw_warning.get("code", "")).strip()
            message = str(raw_warning.get("message", "")).strip()
            if not message:
                continue
            fallback = f"[{code}] {message}" if code else message
            rendered_warnings.append(
                render_file_admission_code(
                    code,
                    declared_format=declared_format,
                    detected_format=detected_format,
                    fallback=fallback,
                )
            )
    if rendered_warnings:
        return " ".join(rendered_warnings)

    warning_code = str(getattr(inspection, "warning_code", ""))
    warning_message = str(getattr(inspection, "warning_message", ""))
    fallback = warning_message or reason_message
    return render_file_admission_message(
        warning_code=warning_code,
        reason_code=reason_code,
        declared_format=declared_format,
        detected_format=detected_format,
        fallback=fallback,
    )


__all__ = [
    "render_file_admission_code",
    "render_file_admission_message",
    "render_file_format_notice",
    "render_file_inspection_message",
]
