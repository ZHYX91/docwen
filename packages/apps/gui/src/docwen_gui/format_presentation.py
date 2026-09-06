"""Single presentation contract for formats shown by the GUI.

Runtime capabilities still decide which routes exist.  This module only owns
how those routes are named and explained, plus UI-only constraints such as the
reliable target set for size-bounded image compression.
"""

from __future__ import annotations

from dataclasses import dataclass

from docwen_core.formats.categories import get_category
from docwen_gui.i18n import t as _t


@dataclass(frozen=True, slots=True)
class FormatPresentation:
    """Stable UI metadata for one canonical format."""

    key: str
    display_name: str
    category: str
    tone: str
    operations: frozenset[str]
    supports_size_limit: bool = False
    allow_same_format_compression: bool = False
    help_key: str = "conversion_panel.format_target_help"
    same_format_disabled_key: str = "conversion_panel.same_format_unavailable"
    size_limit_disabled_key: str = "conversion_panel.image_limit_target_unavailable"


@dataclass(frozen=True, slots=True)
class FormatChoice:
    """One target projected into a selector, including disabled reasons."""

    key: str
    display_name: str
    tone: str
    enabled: bool
    disabled_reason: str
    help_text: str


_TONES: dict[str, str] = {
    "docx": "primary",
    "doc": "info",
    "odt": "success",
    "rtf": "warning",
    "wps": "info",
    "txt": "secondary",
    "xlsx": "primary",
    "xls": "info",
    "ods": "success",
    "csv": "warning",
    "tsv": "warning",
    "et": "info",
    "png": "primary",
    "jpg": "primary",
    "jpeg": "primary",
    "bmp": "info",
    "gif": "success",
    "tif": "warning",
    "tiff": "warning",
    "webp": "danger",
    "heic": "info",
    "heif": "info",
    "pdf": "danger",
    "ofd": "success",
    "xps": "info",
    "md": "secondary",
    "markdown": "secondary",
    "html": "info",
    "htm": "info",
    "mhtml": "info",
    "mht": "info",
    "epub": "success",
    "enex": "warning",
    "ppt": "info",
    "pptx": "primary",
    "ceb": "warning",
}

_DISPLAY_NAMES: dict[str, str] = {
    "jpg": "JPG",
    "jpeg": "JPEG",
    "tif": "TIF",
    "tiff": "TIFF",
    "webp": "WebP",
    "markdown": "Markdown",
}

_SIZE_LIMIT_TARGETS = frozenset({"jpg", "jpeg", "webp"})

SUPPORTED_FORMAT_GROUPS: tuple[tuple[str, str, tuple[str, ...]], ...] = (
    ("file_category.text_short", "Text", ("md", "txt")),
    ("file_category.layout_short", "Layout", ("pdf", "xps", "ofd")),
    ("file_category.document_short", "Doc", ("docx", "doc", "wps", "rtf", "odt")),
    ("file_category.spreadsheet_short", "Sheet", ("xlsx", "xls", "et", "csv", "tsv", "ods")),
    ("file_category.image_short", "Image", ("jpg", "png", "bmp", "gif", "tiff", "heic", "heif", "webp")),
    ("file_category.other_short", "Other", ("html", "mhtml", "enex", "pptx", "ppt", "epub")),
)


def _presentation(key: str) -> FormatPresentation:
    normalized = key.strip().lower()
    supports_size_limit = normalized in _SIZE_LIMIT_TARGETS
    operations = {"conversion"}
    category = get_category(normalized)
    if category in {"document", "spreadsheet", "image"}:
        operations.add("layout-output")
    if supports_size_limit:
        operations.add("size-limit")
    return FormatPresentation(
        key=normalized,
        display_name=_DISPLAY_NAMES.get(normalized, normalized.upper()),
        category=category,
        tone=_TONES.get(normalized, "secondary"),
        operations=frozenset(operations),
        supports_size_limit=supports_size_limit,
        allow_same_format_compression=supports_size_limit,
    )


FORMAT_PRESENTATIONS: dict[str, FormatPresentation] = {key: _presentation(key) for key in _TONES}


def normalize_format(fmt: str) -> str:
    """Normalize aliases used for same-format decisions."""

    normalized = str(fmt or "").strip().lower()
    return {
        "jpg": "jpeg",
        "jpeg": "jpeg",
        "tif": "tiff",
        "tiff": "tiff",
        "heif": "heic",
        "heic": "heic",
        "markdown": "md",
    }.get(normalized, normalized)


def presentation_for(fmt: str) -> FormatPresentation:
    """Return registered metadata, with a safe neutral fallback."""

    key = str(fmt or "").strip().lower()
    return FORMAT_PRESENTATIONS.get(key, _presentation(key))


def format_choice(
    target: str,
    *,
    current_format: str,
    compression_mode: str,
) -> FormatChoice:
    """Project a runtime-backed target into an explained UI choice."""

    presentation = presentation_for(target)
    same_format = normalize_format(target) == normalize_format(current_format)
    limit_mode = presentation.category == "image" and compression_mode == "limit_size"

    disabled_reason = ""
    enabled = True
    if limit_mode and not presentation.supports_size_limit:
        enabled = False
        disabled_reason = _t(
            "conversion_panel.image_limit_target_unavailable",
            "This format cannot reliably meet an exact file-size limit.",
        )
    elif same_format and not (limit_mode and presentation.allow_same_format_compression):
        enabled = False
        disabled_reason = _t(
            "conversion_panel.same_format_unavailable",
            "The source already uses this format.",
        )

    help_text = _t(
        "conversion_panel.format_target_help",
        "Convert to {format}",
        format=presentation.display_name,
    )
    return FormatChoice(
        key=presentation.key,
        display_name=presentation.display_name,
        tone=presentation.tone,
        enabled=enabled,
        disabled_reason=disabled_reason,
        help_text=help_text,
    )


__all__ = [
    "FORMAT_PRESENTATIONS",
    "SUPPORTED_FORMAT_GROUPS",
    "FormatChoice",
    "FormatPresentation",
    "format_choice",
    "normalize_format",
    "presentation_for",
]
