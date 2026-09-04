"""Shared neutral card and internal section-title styles."""

from __future__ import annotations

from .design_tokens import Border, Radius, Typography


def build_panel_card_stylesheet(font_size_preset: str | None = None) -> str:
    """Build the common card hierarchy used by action and conversion panels."""

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
            "    background: transparent;",
            f"    font-size: {Typography.qss(Typography.CAPTION_SIZE, font_size_preset)};",
            "    font-weight: 600;",
            "    padding: 1px 0 7px 0;",
            "    border: none;",
            "    border-bottom: 1px solid palette(midlight);",
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
        ]
    )


__all__ = ["build_panel_card_stylesheet"]
