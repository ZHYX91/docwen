"""操作面板基础样式。"""

from __future__ import annotations

from .control_metrics import button_geometry_qss
from .design_tokens import Border, Sizing, Typography
from .theme_semantics import COLOR_ACCENT_HOVER, COLOR_ACCENT_PRESSED


def build_action_area_stylesheet(font_size_preset: str | None = None) -> str:
    """操作面板基础样式。"""
    return "\n".join(
        [
            "/* docwen-action-panel-foundation */",
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
            "    background-color: palette(highlight);",
            button_geometry_qss(minimum_height=Sizing.ACTION_HEIGHT, minimum_width=Sizing.BUTTON_MIN_WIDTH),
            f"    font-size: {Typography.qss(Typography.BODY_SIZE, font_size_preset)};",
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
            "QWidget#actionAreaRoot QPushButton#actionPrimaryButton:focus {",
            "    border-color: palette(text);",
            "}",
        ]
    )
