"""
格式与类别的统一定义（面向转换编排层）

约定：
- “类别”用于策略查找与路由：markdown/document/spreadsheet/layout/image/unknown
- GUI 内部可能仍使用 'text' 作为“文本选项卡”概念；进入策略层前需归一为 'markdown'
"""

from __future__ import annotations

CATEGORY_MARKDOWN = "markdown"
CATEGORY_DOCUMENT = "document"
CATEGORY_SPREADSHEET = "spreadsheet"
CATEGORY_LAYOUT = "layout"
CATEGORY_IMAGE = "image"
CATEGORY_UNKNOWN = "unknown"


def get_strategy_category_from_format(fmt: str | None) -> str:
    if not fmt:
        return CATEGORY_UNKNOWN
    value = fmt.strip().lower()

    if value in {"md", "markdown", "txt", "text"}:
        return CATEGORY_MARKDOWN

    if value in {"docx", "doc", "rtf", "odt", "wps"}:
        return CATEGORY_DOCUMENT

    if value in {"xlsx", "xls", "et", "ods", "csv"}:
        return CATEGORY_SPREADSHEET

    if value in {"pdf", "ofd", "xps", "caj"}:
        return CATEGORY_LAYOUT

    if value in {"jpg", "jpeg", "png", "gif", "bmp", "tiff", "tif", "webp", "heic", "heif"}:
        return CATEGORY_IMAGE

    return CATEGORY_UNKNOWN


def category_from_actual_format(actual_format: str | None) -> str:
    return get_strategy_category_from_format(actual_format)
