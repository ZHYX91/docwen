"""禁用态语义按钮样式覆盖层。"""

from __future__ import annotations

from ._hex_helper import _hex_to_rgba
from .design_tokens import Border
from .theme_semantics import (
    COLOR_DANGER,
    COLOR_INFO,
    COLOR_PRIMARY,
    COLOR_SECONDARY,
    COLOR_SUCCESS,
    COLOR_WARNING,
    is_dark_theme,
)

_BUTTON_CLASS_COLORS: dict[str, str] = {
    "primary": COLOR_PRIMARY,
    "success": COLOR_SUCCESS,
    "warning": COLOR_WARNING,
    "danger": COLOR_DANGER,
    "info": COLOR_INFO,
    "secondary": COLOR_SECONDARY,
}


def build_disabled_button_stylesheet(theme_name: str) -> str:
    """为语义按钮生成统一的禁用态样式覆盖层。"""
    dark_theme = is_dark_theme(theme_name)
    text_alpha = 220 if dark_theme else 170
    border_alpha = 116 if dark_theme else 96
    background_alpha = 56 if dark_theme else 28
    fallback_text = "rgba(255, 255, 255, 140)" if dark_theme else "rgba(31, 35, 40, 110)"
    fallback_border = "rgba(148, 163, 184, 90)" if dark_theme else "rgba(148, 163, 184, 70)"
    fallback_background = "rgba(148, 163, 184, 34)" if dark_theme else "rgba(148, 163, 184, 20)"

    lines = [
        "/* docwen-disabled-button-overlay */",
        "QPushButton:disabled {",
        f"    color: {fallback_text};",
        f"    border-color: {fallback_border};",
        f"    background-color: {fallback_background};",
        "}",
    ]
    for theme_class, color in _BUTTON_CLASS_COLORS.items():
        lines.extend(
            [
                f'QPushButton[class="{theme_class}"]:disabled {{',
                f"    color: {_hex_to_rgba(color, text_alpha)};",
                f"    border: {Border.THIN}px solid {_hex_to_rgba(color, border_alpha)};",
                f"    background-color: {_hex_to_rgba(color, background_alpha)};",
                "}",
            ]
        )
    return "\n".join(lines)
