"""Paint an icon directly on the widget's current screen/device transform."""

from PySide6.QtGui import QIcon, QPainter, QPaintEvent
from PySide6.QtWidgets import QLabel, QWidget


class IconLabel(QLabel):
    def __init__(self, icon: QIcon, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._icon = icon

    def paintEvent(self, event: QPaintEvent) -> None:
        super().paintEvent(event)
        painter = QPainter(self)
        self._icon.paint(painter, self.contentsRect())
