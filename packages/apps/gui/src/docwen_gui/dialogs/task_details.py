"""Read-only, copyable failure details; viewing never creates another event."""

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QDialog, QDialogButtonBox, QPlainTextEdit, QVBoxLayout, QWidget

from docwen_gui.i18n import t


class TaskDetailsDialog(QDialog):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("taskFailureDetailsDialog")
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        self.setWindowTitle(t("info_area.task_guide_view_failed_details", "Failure details"))
        self.resize(640, 420)
        self.setMinimumSize(300, 220)
        layout = QVBoxLayout(self)
        self.details = QPlainTextEdit(self)
        self.details.setReadOnly(True)
        self.details.setAccessibleName(self.windowTitle())
        layout.addWidget(self.details)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close, self)
        close = buttons.button(QDialogButtonBox.StandardButton.Close)
        if close is not None:
            close.setText(t("common.close", "Close"))
        copy = buttons.addButton(t("common.copy", "Copy"), QDialogButtonBox.ButtonRole.ActionRole)
        copy.clicked.connect(lambda: QApplication.clipboard().setText(self.details.toPlainText()))
        buttons.rejected.connect(self.close)
        layout.addWidget(buttons)

    def set_details(self, text: str) -> None:
        if self.details.toPlainText() != text:
            self.details.setPlainText(text)
