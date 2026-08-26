"""Request-owned OCR presentation policy derived by the runtime."""

from __future__ import annotations

import re
from collections.abc import Mapping
from pathlib import Path
from typing import Any

from docwen_runtime.i18n import load_locale_table

_LOCALE_CODE = re.compile(r"[A-Za-z]{2,3}_[A-Za-z]{2,4}")


def build_ocr_blockquote_title(
    config_snapshot: Mapping[str, Any] | None,
    *,
    requested_locale: object = None,
    locales_dir: Path | None = None,
) -> str:
    """Resolve the OCR blockquote title for one admitted request.

    Explicit per-locale overrides take precedence over the shipped locale
    fallback.  Disabled titles and missing/invalid locale data resolve to an
    empty string rather than leaking a process-global translator into plugins.
    """
    snapshot = config_snapshot if isinstance(config_snapshot, Mapping) else {}
    conversion = _mapping(snapshot.get("conversion"))
    ocr_output = _mapping(conversion.get("ocr_output"))
    if not bool(ocr_output.get("show_blockquote_title", True)):
        return ""

    locale = _resolve_locale(snapshot, requested_locale)
    overrides = _mapping(ocr_output.get("blockquote_title_override_by_locale"))
    override = _plain_title(overrides.get(locale))
    if override:
        return override

    if locales_dir is None:
        from docwen_runtime.resources import ResourceRegistry

        locales_dir = ResourceRegistry.default().locales_dir()
    table = load_locale_table(locales_dir / f"{locale}.toml")
    conversion_table = _mapping(table.get("conversion"))
    ocr_table = _mapping(conversion_table.get("ocr_output"))
    return _plain_title(ocr_table.get("blockquote_prefix"))


def _resolve_locale(snapshot: Mapping[str, Any], requested_locale: object) -> str:
    requested = _normalized_locale(requested_locale)
    if requested:
        return requested
    gui = _mapping(snapshot.get("gui"))
    language = gui.get("language")
    if isinstance(language, Mapping):
        configured = _normalized_locale(language.get("locale"))
        if configured:
            return configured
    configured = _normalized_locale(language)
    if configured:
        return configured
    return "zh_CN"


def _normalized_locale(value: object) -> str:
    """Accept locale codes only, never request-controlled filesystem paths."""
    candidate = value.strip() if isinstance(value, str) else ""
    return candidate if _LOCALE_CODE.fullmatch(candidate) else ""


def _mapping(value: object) -> Mapping[str, Any]:
    return value if isinstance(value, Mapping) else {}


def _plain_title(value: object) -> str:
    """Preserve the locale resource's request-owned Markdown fragment."""
    return str(value or "").strip()


__all__ = ["build_ocr_blockquote_title"]
