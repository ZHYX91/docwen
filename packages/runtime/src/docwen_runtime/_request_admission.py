"""Pure request admission helpers shared by public runtime entry points."""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import replace
from typing import Any

from docwen_core.models.request import ConversionRequest

_MISSING = object()


def _nested_value(values: Mapping[str, Any], *parts: str) -> object:
    current: object = values
    for part in parts:
        if not isinstance(current, Mapping) or part not in current:
            return _MISSING
        current = current[part]
    return current


def _nonblank_text(value: object) -> str | None:
    if not isinstance(value, str):
        return None
    normalized = value.strip()
    return normalized or None


def project_markdown_ocr_options(
    request: ConversionRequest,
    config_snapshot: dict[str, Any],
) -> dict[str, Any]:
    """Fill absent Markdown OCR options from one admitted snapshot.

    Presence, rather than truthiness, defines an explicit request override so
    callers can deliberately pass falsey values. The function is pure and
    never reads live configuration or mutates the caller's request.
    """
    options = dict(request.options)
    if request.target_format != "md" or request.action_name == "process_md_numbering" or not config_snapshot:
        return options

    if "ocr_language" not in options:
        ocr_language = _nonblank_text(_nested_value(config_snapshot, "image", "ocr_language"))
        options["ocr_language"] = ocr_language or "auto"

    if "locale" not in options:
        locale_value = _nested_value(config_snapshot, "gui", "language", "locale")
        locale = _nonblank_text(locale_value)
        options["locale"] = locale or "zh_CN"

    return options


def admit_markdown_ocr_options(
    request: ConversionRequest,
    config_snapshot: dict[str, Any],
) -> ConversionRequest:
    """Return the admitted request without mutating the caller-owned value."""
    options = project_markdown_ocr_options(request, config_snapshot)
    if config_snapshot is request.config_snapshot and options == request.options:
        return request
    return replace(
        request,
        options=options,
        config_snapshot=config_snapshot,
    )
