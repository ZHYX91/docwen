"""An accessible, compact file-location action shared by input and template rows."""

from PySide6.QtCore import QSize, Qt
from PySide6.QtGui import QKeyEvent
from PySide6.QtWidgets import QStyle, QToolButton, QWidget

from docwen_gui.i18n import t
from docwen_gui.styles.design_tokens import Sizing


class LocationButton(QToolButton):
    def __init__(self, parent: QWidget | None = None, *, label: str = "") -> None:
        super().__init__(parent)
        label = label or t("components.template_selector.open_location")
        self.setToolTip(label)
        self.setAccessibleName(label)
        self.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonIconOnly)
        self.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DirOpenIcon))
        self.setIconSize(QSize(20, 20))
        self.setFixedSize(Sizing.CONTROL_HEIGHT, Sizing.CONTROL_HEIGHT)
        self.setAutoRaise(True)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

    def keyPressEvent(self, event: QKeyEvent) -> None:
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            if not event.isAutoRepeat():
                self.click()
            event.accept()
            return
        super().keyPressEvent(event)
