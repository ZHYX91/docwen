"""Shared native checkbox interaction with a scalable vector indicator."""

from PySide6.QtCore import QRectF, Qt
from PySide6.QtGui import QPainter, QPainterPath, QPaintEvent, QPalette, QPen
from PySide6.QtWidgets import QCheckBox, QStyle, QStyleOptionButton, QStyleOptionFocusRect, QStylePainter, QWidget

from docwen_gui.styles.design_tokens import Sizing
from docwen_gui.styles.ui_scale import set_metric


class CheckBox(QCheckBox):
    """Qt owns labels, keyboard, hit testing and accessibility; no bitmap ticks."""

    def __init__(self, text: str = "", parent: QWidget | None = None) -> None:
        super().__init__(text, parent)
        set_metric(self, "setMinimumHeight", Sizing.CONTROL_HEIGHT)

    def paintEvent(self, event: QPaintEvent) -> None:
        del event
        option = QStyleOptionButton()
        self.initStyleOption(option)
        painter = QStylePainter(self)
        label_option = QStyleOptionButton(option)
        label_option.rect = self.style().subElementRect(QStyle.SubElement.SE_CheckBoxContents, option, self)
        painter.drawControl(QStyle.ControlElement.CE_CheckBoxLabel, label_option)
        if option.state & QStyle.StateFlag.State_HasFocus:
            focus = QStyleOptionFocusRect()
            focus.initFrom(self)
            focus.rect = self.style().subElementRect(QStyle.SubElement.SE_CheckBoxFocusRect, option, self)
            painter.drawPrimitive(QStyle.PrimitiveElement.PE_FrameFocusRect, focus)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        indicator = QRectF(self.style().subElementRect(QStyle.SubElement.SE_CheckBoxIndicator, option, self))
        unit = indicator.width() / 18
        rect = indicator.adjusted(unit / 2, unit / 2, -unit / 2, -unit / 2)
        group = QPalette.ColorGroup.Active if self.isEnabled() else QPalette.ColorGroup.Disabled
        active = bool(option.state & (QStyle.StateFlag.State_On | QStyle.StateFlag.State_NoChange))
        hovered = bool(option.state & QStyle.StateFlag.State_MouseOver)
        border = option.palette.color(
            group, QPalette.ColorRole.Highlight if active or hovered else QPalette.ColorRole.Mid
        )
        background = option.palette.color(group, QPalette.ColorRole.Highlight if active else QPalette.ColorRole.Base)
        painter.setPen(QPen(border, unit))
        painter.setBrush(background)
        painter.drawRoundedRect(rect, 4 * unit, 4 * unit)
        if active:
            painter.setPen(
                QPen(
                    option.palette.color(group, QPalette.ColorRole.HighlightedText),
                    1.8 * unit,
                    Qt.PenStyle.SolidLine,
                    Qt.PenCapStyle.RoundCap,
                    Qt.PenJoinStyle.RoundJoin,
                )
            )
            mark = QPainterPath()
            if option.state & QStyle.StateFlag.State_NoChange:
                mark.moveTo(indicator.left() + 4 * unit, indicator.center().y())
                mark.lineTo(indicator.right() - 4 * unit, indicator.center().y())
            else:
                mark.moveTo(indicator.left() + 4 * unit, indicator.top() + 9 * unit)
                mark.lineTo(indicator.left() + 7.3 * unit, indicator.top() + 12.3 * unit)
                mark.lineTo(indicator.left() + 14 * unit, indicator.top() + 5.8 * unit)
            painter.drawPath(mark)
