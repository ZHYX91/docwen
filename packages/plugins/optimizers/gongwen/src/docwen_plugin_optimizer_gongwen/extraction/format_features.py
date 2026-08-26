"""Extract format features (font, size, alignment) from DOCX paragraphs.

Generic extraction functions are re-exported from
``docwen_core.docx_parsing.format_features``.  This module keeps only
gongwen-specific font classification helpers.
"""

from __future__ import annotations

from docwen_core.docx_parsing.format_features import (
    extract_alignment,
    extract_font_info,
    extract_outline_level,
)

__all__ = [
    "extract_alignment",
    "extract_font_info",
    "extract_outline_level",
    "is_body_font",
    "is_heiti_font",
    "is_title_font",
]


def is_title_font(font_name: str) -> bool:
    """Check if a font name matches known official document title fonts."""
    from docwen_plugin_optimizer_gongwen.constants import TITLE_FONTS

    return any(tf in font_name for tf in TITLE_FONTS)


def is_heiti_font(font_name: str) -> bool:
    """Check if a font name matches known 黑体 fonts."""
    from docwen_plugin_optimizer_gongwen.constants import HEITI_FONTS

    return any(hf in font_name for hf in HEITI_FONTS)


def is_body_font(font_name: str) -> bool:
    """Check if a font name matches known body fonts (仿宋)."""
    from docwen_plugin_optimizer_gongwen.constants import BODY_FONTS

    return any(bf in font_name for bf in BODY_FONTS)
