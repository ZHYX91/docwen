"""Compact path labels with complete accessible text and safe tooltips."""

from __future__ import annotations

import html
import unicodedata

from PySide6.QtCore import QEvent, Qt
from PySide6.QtWidgets import QLabel, QSizePolicy, QWidget


def safe_tooltip_text(text: str, *, limit: int = 512) -> str:
    """Render bounded user text literally inside Qt's rich tooltip surface."""

    normalized = str(text).replace("\r\n", "\n").replace("\r", "\n")
    visible = "".join(
        char if char in {"\n", "\t"} or not unicodedata.category(char).startswith("C") else "\ufffd"
        for char in normalized
    )
    truncated = len(visible) > limit
    visible = visible[:limit]
    if truncated:
        visible += "…"
    escaped = html.escape(visible, quote=True).replace("\n", "<br/>").replace("\t", "&#9;")
    return f"<qt>{escaped}</qt>"


class MiddleElidedLabel(QLabel):
    def __init__(self, text: str = "", parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._full_text = ""
        self.setTextFormat(Qt.TextFormat.PlainText)
        self.setWordWrap(False)
        self.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Fixed)
        self.set_full_text(text)

    @property
    def full_text(self) -> str:
        return self._full_text

    def set_full_text(self, text: str) -> None:
        self._full_text = text
        self.setToolTip(safe_tooltip_text(text))
        self.setAccessibleName(text)
        self._refresh_elision()

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._refresh_elision()

    def changeEvent(self, event) -> None:
        super().changeEvent(event)
        if event.type() in {QEvent.Type.FontChange, QEvent.Type.StyleChange}:
            self._refresh_elision()

    def _refresh_elision(self) -> None:
        width = self.contentsRect().width()
        text = self.fontMetrics().elidedText(self._full_text, Qt.TextElideMode.ElideMiddle, width)
        super().setText(text if width > 0 else self._full_text)
