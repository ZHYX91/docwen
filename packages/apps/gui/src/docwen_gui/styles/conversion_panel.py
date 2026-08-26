"""转换面板基础样式。"""

from __future__ import annotations

from ._hex_helper import _hex_to_rgba
from .design_tokens import Border, Radius, Spacing, Typography
from .theme_semantics import (
    COLOR_ACCENT_HOVER,
    COLOR_ACCENT_PRESSED,
    COLOR_SECONDARY,
    COLOR_WARNING,
    get_theme_class_color,
    is_dark_theme,
)


def build_conversion_panel_stylesheet(theme_name: str, font_size_preset: str | None = None) -> str:
    """转换面板基础样式。"""
    dark_theme = is_dark_theme(theme_name)
    hint_color = _hex_to_rgba(COLOR_SECONDARY, 186 if dark_theme else 120)
    detail_hint_color = _hex_to_rgba(COLOR_SECONDARY, 168 if dark_theme else 104)
    warning_hint_color = _hex_to_rgba(COLOR_WARNING, 168 if dark_theme else 128)
    nested_group_border = _hex_to_rgba(COLOR_SECONDARY, 136 if dark_theme else 64)
    nested_group_background = _hex_to_rgba("#1F2937" if dark_theme else "#F8FAFC", 214 if dark_theme else 246)
    nested_group_title = _hex_to_rgba("#E2E8F0" if dark_theme else "#334155", 182 if dark_theme else 154)
    # 分区语义色（现代简约式：标题同色 + 淡色描边/极淡底色，语义沿用旧版 绿=格式转换 红=另存 黄=扩展）
    accent_success = get_theme_class_color("success", theme_name)
    accent_danger = get_theme_class_color("danger", theme_name)
    accent_warning = get_theme_class_color("warning", theme_name)
    accent_border_alpha = 132 if dark_theme else 130
    accent_bg_alpha = 22 if dark_theme else 10
    success_border = _hex_to_rgba(accent_success, accent_border_alpha)
    danger_border = _hex_to_rgba(accent_danger, accent_border_alpha)
    warning_border = _hex_to_rgba(accent_warning, accent_border_alpha)
    success_bg = _hex_to_rgba(accent_success, accent_bg_alpha)
    danger_bg = _hex_to_rgba(accent_danger, accent_bg_alpha)
    warning_bg = _hex_to_rgba(accent_warning, accent_bg_alpha)
    return "\n".join(
        [
            "/* docwen-conversion-panel-foundation */",
            "QWidget#conversionPanelRoot QGroupBox#conversionPrimaryGroup,",
            "QWidget#conversionPanelRoot QGroupBox#conversionSecondaryGroup,",
            "QWidget#conversionPanelRoot QGroupBox#conversionExtraGroup {",
            f"    border-radius: {Radius.LARGE}px;",
            f"    margin-top: {Spacing.MD}px;",
            f"    padding-top: {Spacing.SM}px;",
            "}",
            # -- Section accent tones (特异性高于 panel.py 基础规则) --
            'QWidget#conversionPanelRoot QGroupBox#conversionPrimaryGroup[accentTone="success"] {',
            f"    border: {Border.THIN}px solid {success_border};",
            f"    background-color: {success_bg};",
            "}",
            'QWidget#conversionPanelRoot QGroupBox#conversionPrimaryGroup[accentTone="success"]::title {',
            f"    color: {accent_success};",
            "}",
            'QWidget#conversionPanelRoot QGroupBox#conversionSecondaryGroup[accentTone="danger"] {',
            f"    border: {Border.THIN}px solid {danger_border};",
            f"    background-color: {danger_bg};",
            "}",
            'QWidget#conversionPanelRoot QGroupBox#conversionSecondaryGroup[accentTone="danger"]::title {',
            f"    color: {accent_danger};",
            "}",
            'QWidget#conversionPanelRoot QGroupBox#conversionExtraGroup[accentTone="warning"] {',
            f"    border: {Border.THIN}px solid {warning_border};",
            f"    background-color: {warning_bg};",
            "}",
            'QWidget#conversionPanelRoot QGroupBox#conversionExtraGroup[accentTone="warning"]::title {',
            f"    color: {accent_warning};",
            "}",
            "QWidget#conversionPanelRoot QScrollArea#conversionPanelScrollArea,",
            "QWidget#conversionPanelRoot QWidget#conversionPanelScrollContent {",
            "    border: none;",
            "    background: transparent;",
            "}",
            "QWidget#conversionPanelRoot QGroupBox#conversionPrimaryGroup QVBoxLayout,",
            "QWidget#conversionPanelRoot QGroupBox#conversionSecondaryGroup QVBoxLayout,",
            "QWidget#conversionPanelRoot QGroupBox#conversionExtraGroup QVBoxLayout {",
            f"    spacing: {Spacing.SM}px;",
            "}",
            # -- Subtle divider --
            "QWidget#conversionPanelRoot QFrame#conversionSubtleDivider {",
            f"    color: {nested_group_border};",
            f"    background-color: {nested_group_border};",
            "    min-height: 1px;",
            "    max-height: 1px;",
            "    border: none;",
            "}",
            # -- Nested option group --
            "QGroupBox#conversionOptionGroup {",
            f"    border: {Border.THIN}px solid {nested_group_border};",
            f"    border-radius: {Radius.MEDIUM}px;",
            f"    background-color: {nested_group_background};",
            f"    margin-top: {Spacing.SM}px;",
            f"    padding-top: {Spacing.SM}px;",
            "}",
            "QGroupBox#conversionOptionGroup::title {",
            f"    color: {nested_group_title};",
            f"    font-size: {Typography.qss(Typography.BODY_SIZE, font_size_preset)};",
            "    font-weight: 600;",
            f"    left: {Spacing.SM + 2}px;",
            "    padding: 0 4px;",
            "}",
            "QWidget#conversionPanelRoot QWidget#conversionOptionRow {",
            "    background: transparent;",
            "}",
            "QLabel#conversionOptionLabel {",
            f"    color: {detail_hint_color};",
            f"    font-size: {Typography.qss(Typography.BODY_SIZE, font_size_preset)};",
            "    font-weight: 600;",
            "}",
            # -- Nested group boxes (QGroupBox children of section groups) --
            "QWidget#conversionPanelRoot QGroupBox::title {",
            "    subcontrol-origin: margin;",
            f"    left: {Spacing.MD}px;",
            "    padding: 0 6px;",
            "}",
            "QWidget#conversionPanelRoot QGroupBox#conversionPrimaryGroup QGroupBox,",
            "QWidget#conversionPanelRoot QGroupBox#conversionSecondaryGroup QGroupBox,",
            "QWidget#conversionPanelRoot QGroupBox#conversionExtraGroup QGroupBox {",
            f"    border: {Border.THIN}px solid {nested_group_border};",
            f"    border-radius: {Radius.MEDIUM}px;",
            f"    background-color: {nested_group_background};",
            f"    margin-top: {Spacing.SM}px;",
            f"    padding-top: {Spacing.SM}px;",
            "}",
            "QWidget#conversionPanelRoot QGroupBox#conversionPrimaryGroup QGroupBox::title,",
            "QWidget#conversionPanelRoot QGroupBox#conversionSecondaryGroup QGroupBox::title,",
            "QWidget#conversionPanelRoot QGroupBox#conversionExtraGroup QGroupBox::title {",
            f"    color: {nested_group_title};",
            f"    font-size: {Typography.qss(Typography.BODY_SIZE, font_size_preset)};",
            "    font-weight: 600;",
            f"    left: {Spacing.SM + 2}px;",
            "    padding: 0 4px;",
            "}",
            "QWidget#conversionPanelRoot QLabel#hintLabel {",
            f"    color: {hint_color};",
            f"    font-size: {Typography.qss(Typography.BODY_SIZE, font_size_preset)};",
            "}",
            "QWidget#conversionPanelRoot QLabel#conversionSectionDescription {",
            f"    color: {hint_color};",
            f"    font-size: {Typography.qss(Typography.BODY_SIZE, font_size_preset)};",
            "}",
            "QWidget#conversionPanelRoot QLabel#conversionDetailLabel {",
            f"    color: {detail_hint_color};",
            f"    font-size: {Typography.qss(Typography.CAPTION_SIZE, font_size_preset)};",
            "}",
            "QWidget#conversionPanelRoot QLabel#warningLabel {",
            f"    color: {warning_hint_color};",
            f"    font-size: {Typography.qss(Typography.CAPTION_SIZE, font_size_preset)};",
            "}",
            "QWidget#conversionPanelRoot QWidget#conversionButtonRow {",
            "    background: transparent;",
            "}",
            "QWidget#conversionPanelRoot QPushButton#conversionActionButton {",
            "    color: palette(highlighted-text);",
            f"    border: {Border.THIN}px solid palette(highlight);",
            f"    border-radius: {Radius.MEDIUM}px;",
            "    background-color: palette(highlight);",
            f"    padding: {Spacing.XS}px {Spacing.MD}px;",
            "}",
            "QWidget#conversionPanelRoot QPushButton#conversionActionButton:hover {",
            f"    background-color: {COLOR_ACCENT_HOVER};",
            "}",
            "QWidget#conversionPanelRoot QPushButton#conversionActionButton:pressed {",
            f"    background-color: {COLOR_ACCENT_PRESSED};",
            "}",
            "QWidget#conversionPanelRoot QPushButton#conversionActionButton:disabled {",
            "    color: palette(mid);",
            f"    border: {Border.THIN}px solid palette(midlight);",
            "    background-color: palette(alternate-base);",
            "}",
            "QWidget#generalTransparencyValueRow {",
            "    background: transparent;",
            "}",
            "QDoubleSpinBox#generalTransparencySpinBox {",
            "    min-width: 92px;",
            "    max-width: 92px;",
            "}",
            "QLabel#generalTransparencyPercentLabel {",
            f"    color: {hint_color};",
            f"    font-size: {Typography.qss(Typography.BODY_SIZE, font_size_preset)};",
            "    font-weight: 600;",
            "}",
        ]
    )
