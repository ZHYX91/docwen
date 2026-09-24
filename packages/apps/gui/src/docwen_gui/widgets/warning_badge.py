"""A wrapping warning label using the shared semantic vector icon."""

from PySide6.QtCore import QEvent, QRect, Qt
from PySide6.QtGui import QPainter, QPalette
from PySide6.QtWidgets import QLabel, QWidget

from docwen_gui.styles.ui_scale import dp

from ..resources import load_svg_icon
from ..styles.design_tokens import Spacing


class WarningBadge(QLabel):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWordWrap(True)
        self.setTextFormat(Qt.TextFormat.PlainText)
        self._sync_icon_inset()

    def _sync_icon_inset(self) -> None:
        self.setIndent(max(16, self.fontMetrics().height()) + dp(Spacing.SM))

    def changeEvent(self, event: QEvent) -> None:
        super().changeEvent(event)
        if event.type() in {QEvent.Type.FontChange, QEvent.Type.StyleChange}:
            self._sync_icon_inset()

    def paintEvent(self, event) -> None:
        super().paintEvent(event)
        if not self.text():
            return
        icon = load_svg_icon("warning.svg", color=self.palette().color(QPalette.ColorRole.WindowText).name())
        if icon is None:
            return
        size = self.indent() - dp(Spacing.SM)
        contents = self.contentsRect()
        target = QRect(contents.left(), contents.center().y() - size // 2, size, size)
        painter = QPainter(self)
        icon.paint(painter, target)
