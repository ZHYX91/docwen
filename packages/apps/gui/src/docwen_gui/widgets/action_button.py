"""Execution buttons that retain readable captions in narrow workflow cards."""

from __future__ import annotations

from PySide6.QtCore import QEvent, QRect, QSize, Qt
from PySide6.QtGui import QPalette
from PySide6.QtWidgets import QPushButton, QSizePolicy, QStyle, QStyleOptionButton, QStylePainter

from ..styles.design_tokens import Border, Sizing, Spacing


class ActionButton(QPushButton):
    """Keep native button semantics while wrapping long translated actions."""

    def __init__(self, text: str, parent=None) -> None:
        super().__init__(text, parent)
        policy = QSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        policy.setHeightForWidth(True)
        self.setSizePolicy(policy)
        self.setMinimumHeight(Sizing.ACTION_HEIGHT)

    def minimumSizeHint(self) -> QSize:
        return QSize(Sizing.BUTTON_MIN_WIDTH, Sizing.ACTION_HEIGHT)

    def heightForWidth(self, width: int) -> int:
        horizontal_padding = 2 * (Spacing.MD + Border.THIN)
        vertical_padding = 2 * (Spacing.XS + Border.THIN)
        text_rect = self.fontMetrics().boundingRect(
            QRect(0, 0, max(1, width - horizontal_padding), 100000),
            Qt.TextFlag.TextWordWrap,
            self.text(),
        )
        return max(Sizing.ACTION_HEIGHT, text_rect.height() + vertical_padding)

    def setText(self, text: str) -> None:
        super().setText(text)
        self._sync_height()

    def _sync_height(self) -> None:
        height = self.heightForWidth(self.width())
        if self.minimumHeight() != height:
            self.setMinimumHeight(height)
            self.updateGeometry()

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._sync_height()

    def changeEvent(self, event: QEvent) -> None:
        super().changeEvent(event)
        if event.type() in {QEvent.Type.FontChange, QEvent.Type.StyleChange}:
            self._sync_height()

    def paintEvent(self, event) -> None:
        option = QStyleOptionButton()
        self.initStyleOption(option)
        style = self.style()
        if style is None:
            super().paintEvent(event)
            return
        contents = style.subElementRect(QStyle.SubElement.SE_PushButtonContents, option, self)
        if self.fontMetrics().horizontalAdvance(self.text()) <= contents.width():
            super().paintEvent(event)
            return
        painter = QStylePainter(self)
        option.text = ""
        painter.drawControl(QStyle.ControlElement.CE_PushButton, option)
        painter.drawItemText(
            contents,
            Qt.AlignmentFlag.AlignCenter | Qt.TextFlag.TextWordWrap,
            option.palette,
            self.isEnabled(),
            self.text(),
            QPalette.ColorRole.ButtonText,
        )
