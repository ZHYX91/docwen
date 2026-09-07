"""Native settings checkboxes with readable labels at narrow widths."""

from typing import TYPE_CHECKING

from PySide6.QtCore import QEvent, QSize
from PySide6.QtGui import QTextLayout, QTextOption
from PySide6.QtWidgets import QCheckBox, QSizePolicy, QStyle, QStyleOptionButton, QWidget

from ...styles.design_tokens import Sizing

if TYPE_CHECKING:
    from PySide6.QtWidgets import QCheckBox as _CheckBox
else:
    try:
        from qfluentwidgets import CheckBox as _CheckBox
    except ImportError:
        from PySide6.QtWidgets import QCheckBox as _CheckBox


class SettingsCheckBox(_CheckBox):
    """Preserve native interaction while wrapping the painted label as needed."""

    def __init__(self, text: str, parent: QWidget | None = None) -> None:
        self._source_text = ""
        super().__init__(parent)
        policy = QSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        policy.setHeightForWidth(True)
        self.setSizePolicy(policy)
        self.setText(text)

    def text(self) -> str:
        return self._source_text

    def setText(self, text: str) -> None:
        self._source_text = text
        self.setAccessibleName(text)
        self._reflow_text()

    def _lines(self, width: int) -> list[str]:
        style = self.style()
        inset = style.pixelMetric(QStyle.PixelMetric.PM_IndicatorWidth, None, self)
        inset += style.pixelMetric(QStyle.PixelMetric.PM_CheckBoxLabelSpacing, None, self) + 6
        lines: list[str] = []
        for paragraph in self._source_text.split("\n"):
            layout = QTextLayout(paragraph, self.font())
            option = QTextOption()
            option.setWrapMode(QTextOption.WrapMode.WrapAtWordBoundaryOrAnywhere)
            layout.setTextOption(option)
            encoded = paragraph.encode("utf-16-le")
            layout.beginLayout()
            while (line := layout.createLine()).isValid():
                line.setLineWidth(max(1, width - inset))
                start, end = line.textStart(), line.textStart() + line.textLength()
                lines.append(encoded[start * 2 : end * 2].decode("utf-16-le").rstrip())
            layout.endLayout()
            if not paragraph:
                lines.append("")
        return lines

    def heightForWidth(self, width: int) -> int:
        metrics = self.fontMetrics()
        height = metrics.height() + (max(1, len(self._lines(width))) - 1) * metrics.lineSpacing()
        option = QStyleOptionButton()
        self.initStyleOption(option)
        size = self.style().sizeFromContents(QStyle.ContentsType.CT_CheckBox, option, QSize(0, height), self)
        return max(Sizing.CONTROL_HEIGHT, size.height())

    def minimumSizeHint(self) -> QSize:
        return QSize(0, Sizing.CONTROL_HEIGHT)

    def sizeHint(self) -> QSize:
        # Geometry must depend on the original label, never the last narrow paint.
        option = QStyleOptionButton()
        self.initStyleOption(option)
        option.text = self._source_text
        metrics = self.fontMetrics()
        paragraphs = self._source_text.split("\n")
        size = QSize(
            max((metrics.horizontalAdvance(line) for line in paragraphs), default=0),
            metrics.height() + (len(paragraphs) - 1) * metrics.lineSpacing(),
        )
        size = self.style().sizeFromContents(QStyle.ContentsType.CT_CheckBox, option, size, self)
        return QSize(size.width(), max(Sizing.CONTROL_HEIGHT, size.height()))

    def _reflow_text(self) -> None:
        rendered = "\n".join(self._lines(self.width()))
        if QCheckBox.text(self) != rendered:
            super().setText(rendered)
            self.updateGeometry()

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._reflow_text()

    def changeEvent(self, event) -> None:
        super().changeEvent(event)
        if event.type() in {QEvent.Type.FontChange, QEvent.Type.StyleChange}:
            self._reflow_text()
