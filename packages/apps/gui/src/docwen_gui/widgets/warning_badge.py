"""A wrapping warning label using the shared semantic vector icon."""

from PySide6.QtCore import QEvent, QRect, QSize, Qt
from PySide6.QtGui import QPainter, QPalette
from PySide6.QtWidgets import QWidget

from docwen_gui.styles.ui_scale import dp

from ..resources import load_svg_icon
from ..styles.design_tokens import Spacing
from .panel_card import WrappingLabel


class WarningBadge(WrappingLabel):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__("", parent)
        self.setTextFormat(Qt.TextFormat.PlainText)
        self._sync_icon_inset()

    def _sync_icon_inset(self) -> None:
        self.setIndent(max(16, self.fontMetrics().height()) + dp(Spacing.SM))

    def sizeHint(self) -> QSize:
        # QLabel's wrapped size hint can choose several short lines even when
        # the parent has ample room. An aligned badge then receives that narrow
        # width but the height calculated for the full row, clipping its text.
        # Prefer the unwrapped width; heightForWidth handles genuinely narrow rows.
        margins = self.contentsMargins()
        width = (
            self.fontMetrics().size(Qt.TextFlag.TextExpandTabs, self.text()).width()
            + self.indent()
            + margins.left()
            + margins.right()
            + 2 * self.margin()
            + 2
        )
        return QSize(width, self.heightForWidth(width))

    def heightForWidth(self, width: int) -> int:
        margins = self.contentsMargins()
        text_width = max(1, width - margins.left() - margins.right() - 2 * self.margin() - self.indent())
        flags = Qt.TextFlag.TextExpandTabs
        if self.wordWrap():
            flags |= Qt.TextFlag.TextWordWrap
        text_height = self.fontMetrics().boundingRect(QRect(0, 0, text_width, 100000), flags, self.text()).height()
        icon_height = self.indent() - dp(Spacing.SM)
        return max(text_height, icon_height) + margins.top() + margins.bottom() + 2 * self.margin()

    def sync_wrapped_height(self) -> None:
        height = self.heightForWidth(self.width())
        if self.minimumHeight() != height:
            self.setMinimumHeight(height)
            self.updateGeometry()

    def changeEvent(self, event: QEvent) -> None:
        super().changeEvent(event)
        if event.type() in {QEvent.Type.FontChange, QEvent.Type.StyleChange}:
            self._sync_icon_inset()
            self.sync_wrapped_height()

    def paintEvent(self, event) -> None:
        super().paintEvent(event)
        if not self.text():
            return
        icon = load_svg_icon("warning.svg", color=self.palette().color(QPalette.ColorRole.WindowText).name())
        if icon is None:
            return
        size = self.indent() - dp(Spacing.SM)
        contents = self.contentsRect()
        left = (
            contents.right() - size + 1 if self.layoutDirection() == Qt.LayoutDirection.RightToLeft else contents.left()
        )
        target = QRect(left, contents.center().y() - size // 2, size, size)
        painter = QPainter(self)
        icon.paint(painter, target)
