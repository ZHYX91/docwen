from __future__ import annotations

from collections.abc import Mapping
from types import MappingProxyType

from docwen_core.detection import SUPPORTED_EXTENSION_FORMATS, has_supported_filename_declaration
from docwen_core.formats.categories import get_category

FILE_CATEGORY_ORDER: tuple[str, ...] = ("text", "document", "spreadsheet", "layout", "image", "other")


def _gui_category(format_name: str) -> str:
    """Project a Core format into the GUI's six stable picker groups."""

    if format_name in {"txt", "markdown"}:
        return "text"
    category = get_category(format_name)
    return category if category in {"document", "spreadsheet", "layout", "image"} else "other"


def _build_file_extensions_by_category() -> Mapping[str, tuple[str, ...]]:
    groups: dict[str, list[str]] = {category: [] for category in FILE_CATEGORY_ORDER}
    for extension, format_name in SUPPORTED_EXTENSION_FORMATS.items():
        groups[_gui_category(format_name)].append(extension)
    return MappingProxyType({category: tuple(groups[category]) for category in FILE_CATEGORY_ORDER})


FILE_EXTENSIONS_BY_CATEGORY: Mapping[str, tuple[str, ...]] = _build_file_extensions_by_category()


def is_supported_file(file_path: str) -> bool:
    """Delegate filename-declaration support to the Core registry."""

    return has_supported_filename_declaration(file_path)
