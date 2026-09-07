"""状态栏样式。"""

from __future__ import annotations

from ._hex_helper import _hex_to_rgba
from .control_metrics import button_geometry_qss
from .design_tokens import Border, Radius, Sizing, Typography
from .theme_semantics import (
    COLOR_INFO,
    COLOR_SECONDARY,
    get_theme_class_color,
    is_dark_theme,
)


def build_info_area_stylesheet(theme_name: str, font_size_preset: str | None = None) -> str:
    """状态栏样式统一收口到全局主题中心。"""
    dark_theme = is_dark_theme(theme_name)
    section_background = _hex_to_rgba("#111827" if dark_theme else "#FFFFFF", 244 if dark_theme else 255)
    muted_text = _hex_to_rgba("#E2E8F0" if dark_theme else "#0F172A", 212 if dark_theme else 190)
    action_shell_border = _hex_to_rgba(COLOR_SECONDARY, 116 if dark_theme else 72)
    action_shell_background = _hex_to_rgba("#020617" if dark_theme else "#FFFFFF", 210 if dark_theme else 250)
    info_accent = _hex_to_rgba(COLOR_INFO, 192 if dark_theme else 148)
    activity_accent = get_theme_class_color("info", theme_name)
    warning_accent = get_theme_class_color("warning", theme_name)
    return "\n".join(
        [
            "/* docwen-status-bar-foundation */",
            "QToolButton#outputFileCount {",
            f"    border: {Border.THIN}px solid palette(mid);",
            f"    border-radius: {Sizing.CONTROL_HEIGHT // 2}px;",
            "    padding: 0 8px; background: palette(alternate-base); color: palette(text);",
            "}",
            "QToolButton#outputFileCount:hover, QToolButton#outputFileCount:focus {",
            "    border-color: palette(highlight);",
            "}",
            "QWidget#infoArea {",
            "    background: transparent;",
            "}",
            "QLabel#infoOutputDestination {",
            f"    color: {muted_text};",
            f"    font-size: {Typography.qss(Typography.CAPTION_SIZE, font_size_preset)};",
            "    padding: 0;",
            "}",
            "QPushButton#infoStatusSummary {",
            "    border: none; padding: 0; min-width: 0; min-height: 0;",
            f"    color: {muted_text};",
            "    font-weight: 500;",
            "    background: transparent;",
            "}",
            'QPushButton#infoStatusSummary[interactiveText="true"] {',
            f"    color: {info_accent};",
            "}",
            'QPushButton#infoStatusSummary[interactiveTextHovered="true"] {',
            "    color: palette(highlight);",
            "}",
            "QPushButton#infoStatusSummary:focus {",
            f"    border: {Border.THIN}px solid palette(highlight);",
            f"    border-radius: {Radius.SMALL}px;",
            "    padding: 0 2px;",
            "}",
            "QPushButton#infoStatusSummary QLabel {",
            f"    color: {muted_text}; background: transparent; border: none;",
            "}",
            'QPushButton#infoStatusSummary[interactiveText="true"] QLabel {',
            "    color: palette(highlight);",
            "}",
            "QLabel#infoNotification {",
            f"    color: {muted_text}; padding: 6px;",
            "}",
            "QProgressBar#infoTaskProgress { border: none; background: palette(midlight); }",
            "QProgressBar#infoTaskProgress::chunk { background: palette(highlight); }",
            "QWidget#infoStatusGuideRow,",
            "QWidget#infoStatusGuideActions {",
            "    background: transparent;",
            "}",
            "QPushButton#infoStatusGuideButton, QPushButton#secondaryActionButton {",
            f"    border: {Border.THIN}px solid {action_shell_border};",
            f"    border-radius: {Radius.MEDIUM}px;",
            f"    background: {action_shell_background};",
            button_geometry_qss(minimum_width=Sizing.BUTTON_MIN_WIDTH),
            "    font-weight: 500;",
            "}",
            "QPushButton#infoStatusGuideButton:hover {",
            f"    background: {section_background};",
            "}",
            "QPushButton#infoActivityButton {",
            button_geometry_qss(),
            f"    border: {Border.THIN}px solid {action_shell_border};",
            f"    border-radius: {Sizing.CONTROL_HEIGHT // 2}px;",
            f"    background: {_hex_to_rgba(activity_accent, 22 if dark_theme else 12)};",
            "    color: palette(text); font-weight: 500;",
            "}",
            "QPushButton#infoActivityButton:hover, QPushButton#infoActivityButton:focus {",
            f"    border-color: {activity_accent};",
            f"    background: {_hex_to_rgba(activity_accent, 40 if dark_theme else 24)};",
            "}",
            "QPushButton#infoActivityButton:pressed {",
            f"    background: {_hex_to_rgba(activity_accent, 60 if dark_theme else 38)};",
            "}",
            'QPushButton#infoActivityButton[hasFailures="true"] {',
            f"    border-color: {warning_accent};",
            f"    background: {_hex_to_rgba(warning_accent, 30 if dark_theme else 20)};",
            "}",
            'QPushButton#infoActivityButton[hasFailures="true"]:hover, QPushButton#infoActivityButton[hasFailures="true"]:focus {',
            f"    background: {_hex_to_rgba(warning_accent, 52 if dark_theme else 36)};",
            "    border-color: palette(highlight);",
            "}",
        ]
    )
