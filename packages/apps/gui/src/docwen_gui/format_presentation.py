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
    swatch_colors: tuple[str, str] | None
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
    enabled: bool
    disabled_reason: str
    help_text: str


# Light/dark swatches identify file families, independently of task status colors.
# Names remain the primary identifiers; aliases share colors. Related formats use
# nearby hues with visible differences in lightness and saturation.
_SWATCH_COLORS: dict[str, tuple[str, str]] = {
    # Documents: blue, with a softer blue for legacy DOC.
    "docx": ("#2563EB", "#60A5FA"),
    "doc": ("#6586B5", "#B0C7E3"),
    "odt": ("#167F98", "#56B9CA"),
    "rtf": ("#5F879C", "#94B8C9"),
    "wps": ("#5962BA", "#9DABF2"),
    # Spreadsheets: green/teal.
    "xlsx": ("#16824D", "#4DD49A"),
    "xls": ("#64866F", "#A5C9AD"),
    "ods": ("#087D79", "#51C7BE"),
    "csv": ("#477D35", "#A3CA75"),
    "tsv": ("#747B31", "#C0C982"),
    "et": ("#3C8273", "#8BCAB8"),
    # Images: purple/magenta. Equivalent extensions keep the same identity.
    "png": ("#8050C3", "#BC96ED"),
    "jpg": ("#A34F91", "#E3A0CB"),
    "jpeg": ("#A34F91", "#E3A0CB"),
    "bmp": ("#857493", "#BEAFCD"),
    "gif": ("#6446A6", "#A99BEB"),
    "tif": ("#6F619F", "#B0AED6"),
    "tiff": ("#6F619F", "#B0AED6"),
    "webp": ("#9C42BA", "#D88AE7"),
    "heic": ("#875782", "#CAA0C7"),
    "heif": ("#875782", "#CAA0C7"),
    # Fixed layout: red/rose; presentations: orange.
    "pdf": ("#CB4054", "#F58B98"),
    "ofd": ("#AA5B79", "#D9A1B6"),
    "xps": ("#A2665C", "#DCB0A5"),
    "ceb": ("#9B5368", "#D39CA8"),
    "pptx": ("#BD5A21", "#F5A46D"),
    "ppt": ("#A67550", "#D9BA99"),
    # Plain/Markdown text is neutral; web and ebook formats use amber.
    "txt": ("#7B8491", "#B3BDCB"),
    "md": ("#475569", "#CBD5E1"),
    "markdown": ("#475569", "#CBD5E1"),
    "html": ("#A16D16", "#E9BE68"),
    "htm": ("#A16D16", "#E9BE68"),
    "mhtml": ("#8B774E", "#D6C299"),
    "mht": ("#8B774E", "#D6C299"),
    "epub": ("#8D7630", "#D0BF77"),
    "enex": ("#9A7042", "#DDB88B"),
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
        swatch_colors=_SWATCH_COLORS.get(normalized),
        operations=frozenset(operations),
        supports_size_limit=supports_size_limit,
        allow_same_format_compression=supports_size_limit,
    )


FORMAT_PRESENTATIONS: dict[str, FormatPresentation] = {key: _presentation(key) for key in _SWATCH_COLORS}


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
