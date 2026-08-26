"""About dialog stylesheet.

Provides visual styling for the AboutDialog widget, including hero card,
tool list grid, disclaimer card, and close button.
"""

from __future__ import annotations

from .design_tokens import Border, Radius, Typography


def build_about_dialog_stylesheet(font_size_preset: str | None = None) -> str:
    """Return the about-dialog CSS fragment."""
    return "\n".join(
        [
            "/* docwen-about-dialog-foundation */",
            "QDialog#aboutDialog QScrollArea#aboutScrollArea {",
            "    border: none;",
            "    background: transparent;",
            "}",
            "QDialog#aboutDialog QLabel#aboutTitle {",
            f"    font-size: {Typography.qss(Typography.DIALOG_TITLE_SIZE, font_size_preset)};",
            "    font-weight: 700;",
            "}",
            "QDialog#aboutDialog QLabel#aboutSubtitle,",
            "QDialog#aboutDialog QLabel#aboutVersion,",
            "QDialog#aboutDialog QLabel#aboutMeta,",
            "QDialog#aboutDialog QLabel#aboutUpdateNotice {",
            "    color: palette(text);",
            "}",
            "QDialog#aboutDialog QLabel#aboutUpdateNotice {",
            f"    font-size: {Typography.qss(Typography.CAPTION_SIZE, font_size_preset)};",
            "}",
            "QDialog#aboutDialog QLabel#aboutGroupTitle {",
            f"    font-size: {Typography.qss(Typography.CARD_TITLE_SIZE, font_size_preset)};",
            "    font-weight: 600;",
            "}",
            "QDialog#aboutDialog QWidget#aboutHeroCard,",
            "QDialog#aboutDialog QWidget#aboutGroup {",
            f"    border: {Border.THIN}px solid palette(mid);",
            f"    border-radius: {Radius.DEFAULT}px;",
            "    background-color: palette(base);",
            "}",
            "QDialog#aboutDialog QWidget#aboutGroup {",
            "    margin-top: 12px;",
            "    padding-top: 8px;",
            "}",
            "QDialog#aboutDialog QToolButton#aboutToolInfoButton {",
            "    border: none;",
            "    background: transparent;",
            "    padding: 0;",
            "}",
            "QDialog#aboutDialog QToolButton#aboutToolInfoButton:hover {",
            "    background: palette(midlight);",
            f"    border-radius: {Radius.SMALL}px;",
            "}",
            "QDialog#aboutDialog QPushButton#aboutCloseButton {",
            "    min-height: 32px;",
            "}",
        ]
    )
