"""Shared template rows: source badge, elided name, and optional default marker."""

from __future__ import annotations

from PySide6.QtCore import QModelIndex, QRect, QSize, Qt
from PySide6.QtGui import QColor, QFont, QFontMetrics, QPainter, QPalette
from PySide6.QtWidgets import QApplication, QStyle, QStyledItemDelegate, QStyleOptionViewItem

from ..i18n import t

SOURCE_ROLE = Qt.ItemDataRole.UserRole + 20
CUSTOM_ROLE = Qt.ItemDataRole.UserRole + 21
DEFAULT_ROLE = Qt.ItemDataRole.UserRole + 22


class TemplateItemDelegate(QStyledItemDelegate):
    """Draw badges without adding child widgets that intercept selection or dragging."""

    def __init__(self, parent=None, *, draggable: bool = False) -> None:
        super().__init__(parent)
        self.draggable = draggable

    def sizeHint(self, option, index) -> QSize:
        return QSize(0, max(40, option.fontMetrics.height() + 18))

    def paint(self, painter: QPainter, option: QStyleOptionViewItem, index: QModelIndex) -> None:
        source = index.data(SOURCE_ROLE)
        if not source:
            super().paint(painter, option, index)
            return
        row = QStyleOptionViewItem(option)
        self.initStyleOption(row, index)
        name = row.text
        row.text = ""
        style = row.widget.style() if row.widget else QApplication.style()
        painter.save()
        if self.draggable:
            row.rect.adjust(16, 0, 0, 0)
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(row.palette.color(QPalette.ColorRole.Mid))
            for x in (6, 10):
                for y in (-4, 0, 4):
                    painter.drawEllipse(option.rect.left() + x, option.rect.center().y() + y, 2, 2)
        style.drawControl(QStyle.ControlElement.CE_ItemViewItem, row, painter, row.widget)
        content = style.subElementRect(QStyle.SubElement.SE_ItemViewItemText, row, row.widget).adjusted(4, 0, -8, 0)
        badge_font = QFont(row.font)
        badge_font.setPointSizeF(max(8, badge_font.pointSizeF() * 0.82))
        metrics = QFontMetrics(badge_font)
        width = (
            max(
                metrics.horizontalAdvance(t("settings.templates.builtin")),
                metrics.horizontalAdvance(t("settings.templates.custom")),
            )
            + 16
        )
        height = metrics.height() + 6
        badge = QRect(content.left(), content.center().y() - height // 2, width, height)
        ink = row.palette.color(QPalette.ColorRole.Text)
        background = QColor(row.palette.color(QPalette.ColorRole.Highlight) if index.data(CUSTOM_ROLE) else ink)
        background.setAlpha(32 if index.data(CUSTOM_ROLE) else 18)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(row.palette.color(QPalette.ColorRole.Base))
        painter.drawRoundedRect(badge, height / 2, height / 2)
        painter.setBrush(background)
        painter.drawRoundedRect(badge, height / 2, height / 2)
        painter.setPen(ink)
        painter.setFont(badge_font)
        painter.drawText(badge, Qt.AlignmentFlag.AlignCenter, str(source))
        content.setLeft(badge.right() + 9)
        selected = bool(row.state & QStyle.StateFlag.State_Selected)
        text_ink = row.palette.color(QPalette.ColorRole.HighlightedText) if selected and not self.draggable else ink
        painter.setPen(text_ink)
        if index.data(DEFAULT_ROLE):
            label = t("settings.templates.default")
            marker_width = metrics.horizontalAdvance(label) + 8
            marker = QRect(content.right() - marker_width, content.top(), marker_width, content.height())
            painter.drawText(marker, Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignRight, label)
            content.setRight(marker.left() - 8)
        painter.setFont(row.font)
        checked = index.data(Qt.ItemDataRole.CheckStateRole)
        muted = checked == Qt.CheckState.Unchecked.value
        painter.setPen(row.palette.color(QPalette.ColorRole.Mid) if muted and not selected else text_ink)
        painter.drawText(
            content,
            Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft,
            row.fontMetrics.elidedText(name, Qt.TextElideMode.ElideMiddle, max(0, content.width())),
        )
        painter.restore()

    def editorEvent(self, event, model, option, index) -> bool:
        adjusted = QStyleOptionViewItem(option)
        if self.draggable:
            adjusted.rect.adjust(16, 0, 0, 0)
        return super().editorEvent(event, model, adjusted, index)
