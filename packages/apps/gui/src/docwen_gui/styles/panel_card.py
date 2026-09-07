"""Shared semantic card and internal section-title styles."""

from __future__ import annotations

from .design_tokens import Border, Radius, Spacing, Typography
from .theme_semantics import ThemeClass, get_card_colors


def build_panel_card_stylesheet(font_size_preset: str | None = None, *, theme_name: str = "light") -> str:
    """Build the common card hierarchy used by action and conversion panels."""

    tones: tuple[ThemeClass, ...] = ("primary", "info", "success", "warning", "danger", "secondary")
    tone_rules = []
    for tone in tones:
        surface, text = get_card_colors(tone, theme_name)
        tone_rules.extend(
            [
                f'QFrame[panelLevel="card"][panelTone="{tone}"] {{ border-color: {surface}; }}',
                f'QLabel#panelCardTitle[panelTone="{tone}"] {{ background-color: {surface}; color: {text}; }}',
            ]
        )
    return "\n".join(
        [
            'QFrame[panelLevel="card"] {',
            f"    border: {Border.THIN}px solid palette(midlight);",
            f"    border-radius: {Radius.LARGE}px;",
            "    background-color: palette(base);",
            "}",
            'QFrame[panelLevel="section"] {',
            "    border: none;",
            "    background: transparent;",
            "}",
            "QWidget#panelCardContent {",
            "    border: none;",
            "    background: transparent;",
            "}",
            "QLabel#panelCardTitle {",
            "    color: palette(text);",
            "    background: palette(alternate-base);",
            f"    font-size: {Typography.qss(Typography.CARD_TITLE_SIZE, font_size_preset)};",
            "    font-weight: 600;",
            f"    padding: {Spacing.MD}px {Spacing.CARD_PADDING}px;",
            "    border: none;",
            f"    border-top-left-radius: {Radius.LARGE - Border.THIN}px;",
            f"    border-top-right-radius: {Radius.LARGE - Border.THIN}px;",
            "}",
            'QLabel#panelCardTitle[contentCollapsed="true"] {',
            f"    border-bottom-left-radius: {Radius.LARGE - Border.THIN}px;",
            f"    border-bottom-right-radius: {Radius.LARGE - Border.THIN}px;",
            "}",
            "QLabel#panelSectionTitle {",
            "    color: palette(text);",
            "    background: transparent;",
            f"    font-size: {Typography.qss(Typography.BODY_SIZE, font_size_preset)};",
            "    font-weight: 600;",
            "    border: none;",
            "}",
            "QFrame#panelFormRow,",
            "QFrame#panelChoiceGroup,",
            "QFrame#panelActionFooter,",
            "QFrame#taskActivityList {",
            "    border: none;",
            "    background: transparent;",
            "}",
            "QLabel#panelFormLabel {",
            "    color: palette(text);",
            f"    font-size: {Typography.qss(Typography.BODY_SIZE, font_size_preset)};",
            "}",
            "QFrame#panelInlineNotice {",
            f"    border: {Border.THIN}px solid palette(midlight);",
            f"    border-radius: {Radius.MEDIUM}px;",
            "    background-color: palette(alternate-base);",
            "}",
            'QFrame#panelInlineNotice[noticeTone="warning"] {',
            "    border-color: palette(highlight);",
            "}",
            "QFrame#panelInlineNotice QLabel {",
            "    border: none;",
            "    background: transparent;",
            f"    font-size: {Typography.qss(Typography.CAPTION_SIZE, font_size_preset)};",
            "}",
            *tone_rules,
        ]
    )


__all__ = ["build_panel_card_stylesheet"]
