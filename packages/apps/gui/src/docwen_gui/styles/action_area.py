"""操作面板基础样式。"""

from __future__ import annotations

from .design_tokens import Border, Radius, Spacing, Typography
from .theme_semantics import COLOR_ACCENT_HOVER, COLOR_ACCENT_PRESSED


def build_action_area_stylesheet(font_size_preset: str | None = None) -> str:
    """操作面板基础样式。"""
    return "\n".join(
        [
            "/* docwen-action-panel-foundation */",
            "QFrame#actionContentCard, QFrame#actionCancelCard {",
            f"    border: {Border.THIN}px solid palette(midlight);",
            f"    border-radius: {Radius.LARGE}px;",
            "    background-color: palette(base);",
            "}",
            "QLabel#actionPanelTitle {",
            "    color: palette(text);",
            f"    font-size: {Typography.qss(Typography.SECTION_TITLE_SIZE, font_size_preset)};",
            "    font-weight: 600;",
            "}",
            "QLabel#actionPanelSubtitle {",
            "    color: palette(mid);",
            f"    font-size: {Typography.qss(Typography.CAPTION_SIZE, font_size_preset)};",
            "}",
            "QWidget#actionAreaRoot QWidget#actionOptionRow {",
            "    background: transparent;",
            "}",
            "QWidget#actionAreaRoot QPushButton#actionPrimaryButton {",
            "    color: palette(highlighted-text);",
            f"    border: {Border.THIN}px solid palette(highlight);",
            f"    border-radius: {Radius.MEDIUM}px;",
            "    background-color: palette(highlight);",
            f"    padding: {Spacing.XS}px {Spacing.MD}px;",
            "}",
            "QWidget#actionAreaRoot QPushButton#actionPrimaryButton:hover {",
            f"    background-color: {COLOR_ACCENT_HOVER};",
            "}",
            "QWidget#actionAreaRoot QPushButton#actionPrimaryButton:pressed {",
            f"    background-color: {COLOR_ACCENT_PRESSED};",
            "}",
            "QWidget#actionAreaRoot QPushButton#actionPrimaryButton:disabled {",
            "    color: palette(mid);",
            f"    border: {Border.THIN}px solid palette(midlight);",
            "    background-color: palette(alternate-base);",
            "}",
        ]
    )
