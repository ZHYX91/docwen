"""状态栏样式。"""

from __future__ import annotations

from ._hex_helper import _hex_to_rgba
from .design_tokens import Border, Radius, Typography
from .theme_semantics import (
    COLOR_DANGER,
    COLOR_INFO,
    COLOR_SECONDARY,
    COLOR_SUCCESS,
    COLOR_WARNING,
    is_dark_theme,
)


def build_info_area_stylesheet(theme_name: str, font_size_preset: str | None = None) -> str:
    """状态栏样式统一收口到全局主题中心。"""
    dark_theme = is_dark_theme(theme_name)
    section_background = _hex_to_rgba("#111827" if dark_theme else "#FFFFFF", 244 if dark_theme else 255)
    divider_color = _hex_to_rgba(COLOR_SECONDARY, 88 if dark_theme else 46)
    meta_text = _hex_to_rgba("#CBD5E1" if dark_theme else "#475569", 176 if dark_theme else 140)
    muted_text = _hex_to_rgba("#E2E8F0" if dark_theme else "#0F172A", 212 if dark_theme else 190)
    action_shell_border = _hex_to_rgba(COLOR_SECONDARY, 116 if dark_theme else 72)
    action_shell_background = _hex_to_rgba("#020617" if dark_theme else "#FFFFFF", 210 if dark_theme else 250)
    info_accent = _hex_to_rgba(COLOR_INFO, 192 if dark_theme else 148)
    history_info_background = _hex_to_rgba(COLOR_INFO, 18 if dark_theme else 10)
    history_success_background = _hex_to_rgba(COLOR_SUCCESS, 18 if dark_theme else 10)
    history_warning_background = _hex_to_rgba(COLOR_WARNING, 18 if dark_theme else 10)
    history_danger_background = _hex_to_rgba(COLOR_DANGER, 18 if dark_theme else 10)
    return "\n".join(
        [
            "/* docwen-status-bar-foundation */",
            "QWidget#infoArea {",
            "    background: transparent;",
            "}",
            "QWidget#infoHistoryEmptyState {",
            "    background: transparent;",
            "}",
            "QLabel#infoHistoryEmptyTitle {",
            f"    color: {muted_text};",
            "    background: transparent;",
            f"    font-size: {Typography.qss(Typography.EMPHASIS_TITLE_SIZE, font_size_preset)};",
            "    font-weight: 600;",
            "}",
            "QLabel#infoHistoryEmptyCaption {",
            f"    color: {meta_text};",
            "    background: transparent;",
            f"    font-size: {Typography.qss(Typography.BODY_SIZE, font_size_preset)};",
            "}",
            "QFrame#infoSectionDivider {",
            "    border: none;",
            f"    border-top: {Border.THIN}px solid {divider_color};",
            "    max-height: 1px;",
            "}",
            "QLabel#statusTimestamp {",
            f"    color: {meta_text};",
            "    background: transparent;",
            f"    font-size: {Typography.qss(Typography.BODY_SIZE, font_size_preset)};",
            "}",
            "QLabel#infoStatusMeta {",
            f"    color: {meta_text};",
            "    background: transparent;",
            f"    font-size: {Typography.qss(Typography.BODY_SIZE, font_size_preset)};",
            "    font-weight: 600;",
            "}",
            "QLabel#infoStatusSummary {",
            f"    color: {muted_text};",
            "    font-weight: 500;",
            "    background: transparent;",
            "}",
            'QLabel#infoStatusSummary[interactiveText="true"] {',
            f"    color: {info_accent};",
            "}",
            'QLabel#infoStatusSummary[interactiveTextHovered="true"] {',
            "    color: palette(highlight);",
            "}",
            "QLabel#infoStatusSummary:focus {",
            f"    border: {Border.THIN}px solid palette(highlight);",
            f"    border-radius: {Radius.SMALL}px;",
            "    padding: 0 2px;",
            "}",
            "QWidget#infoStatusGuideRow,",
            "QWidget#infoStatusGuideActions {",
            "    background: transparent;",
            "}",
            "QPushButton#infoStatusGuideButton {",
            f"    border: {Border.THIN}px solid {action_shell_border};",
            f"    border-radius: {Radius.MEDIUM}px;",
            f"    background: {action_shell_background};",
            "    padding: 3px 10px;",
            "    font-weight: 500;",
            "}",
            'QPushButton#infoStatusGuideButton[guideActionPriority="primary"] {',
            f"    border-color: {info_accent};",
            "    font-weight: 600;",
            "}",
            "QPushButton#infoStatusGuideButton:hover {",
            f"    background: {section_background};",
            "}",
            "QWidget#infoHistoryRow {",
            "    background: transparent;",
            "}",
            'QWidget#infoHistoryRow[infoStatusTone="info"] {',
            f"    background: {history_info_background};",
            "}",
            'QWidget#infoHistoryRow[infoStatusTone="success"] {',
            f"    background: {history_success_background};",
            "}",
            'QWidget#infoHistoryRow[infoStatusTone="warning"] {',
            f"    background: {history_warning_background};",
            "}",
            'QWidget#infoHistoryRow[infoStatusTone="danger"] {',
            f"    background: {history_danger_background};",
            "}",
            "QWidget#infoHistoryMeta,",
            "QLabel#infoHistoryText,",
            "QScrollArea#infoHistoryScrollArea,",
            "QScrollArea#infoHistoryScrollArea > QWidget > QWidget {",
            "    background: transparent;",
            "}",
            "QScrollArea#infoHistoryScrollArea {",
            "    border: none;",
            "}",
            "QToolButton#statusLocationButton {",
            "    border: none;",
            "    background: transparent;",
            "    padding: 0px;",
            "}",
            "QToolButton#statusLocationButton:hover {",
            "    background: transparent;",
            "}",
        ]
    )
