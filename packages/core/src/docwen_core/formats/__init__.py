"""Format categories, media types, and route registry."""

from docwen_core.formats.categories import (
    CATEGORY_DOCUMENT,
    CATEGORY_IMAGE,
    CATEGORY_LAYOUT,
    CATEGORY_MARKDOWN,
    CATEGORY_MARKUP,
    CATEGORY_OTHER,
    CATEGORY_PRESENTATION,
    CATEGORY_SPREADSHEET,
    FORMAT_CATEGORY,
    FORMAT_MEDIA_TYPE,
    get_category,
    get_media_type,
)
from docwen_core.formats.routes import RouteEntry, RouteRegistry

__all__ = [
    "CATEGORY_DOCUMENT",
    "CATEGORY_IMAGE",
    "CATEGORY_LAYOUT",
    "CATEGORY_MARKDOWN",
    "CATEGORY_MARKUP",
    "CATEGORY_OTHER",
    "CATEGORY_PRESENTATION",
    "CATEGORY_SPREADSHEET",
    "FORMAT_CATEGORY",
    "FORMAT_MEDIA_TYPE",
    "RouteEntry",
    "RouteRegistry",
    "get_category",
    "get_media_type",
]
