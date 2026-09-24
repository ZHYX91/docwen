"""Execution buttons that retain readable captions in narrow workflow cards."""

from __future__ import annotations

from PySide6.QtCore import QEvent, QRect, QSize, Qt
from PySide6.QtGui import QIcon, QPalette
from PySide6.QtWidgets import QPushButton, QSizePolicy, QStyle, QStyleOptionButton, QStylePainter

from docwen_gui.styles.ui_scale import dp, set_metric

from ..styles.design_tokens import Border, Sizing, Spacing


class ActionButton(QPushButton):
    """Keep native button semantics while wrapping long translated actions."""

    def __init__(self, text: str, parent=None) -> None:
        super().__init__(text, parent)
        policy = QSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        policy.setHeightForWidth(True)
        self.setSizePolicy(policy)
        set_metric(self, "setMinimumHeight", Sizing.ACTION_HEIGHT)

    def minimumSizeHint(self) -> QSize:
        return QSize(dp(Sizing.BUTTON_MIN_WIDTH), dp(Sizing.ACTION_HEIGHT))

    def heightForWidth(self, width: int) -> int:
        horizontal_padding = 2 * (dp(Spacing.MD) + dp(Border.THIN))
        if not self.icon().isNull():
            horizontal_padding += self.iconSize().width() + dp(Spacing.SM)
        vertical_padding = 2 * (dp(Spacing.XS) + dp(Border.THIN))
        text_rect = self.fontMetrics().boundingRect(
            QRect(0, 0, max(1, width - horizontal_padding), 100000),
            Qt.TextFlag.TextWordWrap,
            self.text(),
        )
        icon_height = 0 if self.icon().isNull() else self.iconSize().height()
        return max(dp(Sizing.ACTION_HEIGHT), max(text_rect.height(), icon_height) + vertical_padding)

    def setIcon(self, icon: QIcon) -> None:
        super().setIcon(icon)
        self._sync_height()

    def setIconSize(self, size: QSize) -> None:
        super().setIconSize(size)
        self._sync_height()

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
        icon_width = self.iconSize().width() + dp(Spacing.SM) if not self.icon().isNull() else 0
        if self.fontMetrics().horizontalAdvance(self.text()) + icon_width <= contents.width():
            super().paintEvent(event)
            return
        painter = QStylePainter(self)
        option.text = ""
        option.icon = QIcon()
        painter.drawControl(QStyle.ControlElement.CE_PushButton, option)
        if icon_width:
            icon_rect = QRect(
                contents.left(),
                contents.center().y() - self.iconSize().height() // 2,
                self.iconSize().width(),
                self.iconSize().height(),
            )
            icon_rect = QStyle.visualRect(self.layoutDirection(), contents, icon_rect)
            mode = QIcon.Mode.Normal if self.isEnabled() else QIcon.Mode.Disabled
            self.icon().paint(painter, icon_rect, Qt.AlignmentFlag.AlignCenter, mode)
            text_rect = contents.adjusted(icon_width, 0, 0, 0)
            contents = QStyle.visualRect(self.layoutDirection(), contents, text_rect)
        painter.drawItemText(
            contents,
            Qt.AlignmentFlag.AlignCenter | Qt.TextFlag.TextWordWrap,
            option.palette,
            self.isEnabled(),
            self.text(),
            QPalette.ColorRole.ButtonText,
        )
