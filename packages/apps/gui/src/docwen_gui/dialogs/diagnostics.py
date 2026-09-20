"""Local task details and a separate, exact preview of copyable diagnostics."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QDialogButtonBox,
    QLabel,
    QPlainTextEdit,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from docwen_gui.diagnostics import DiagnosticSummary
from docwen_gui.i18n import t
from docwen_gui.styles.design_tokens import Spacing
from docwen_gui.styles.theme_semantics import apply_theme_class


class DiagnosticView(QTabWidget):
    """Keep readable local context out of the explicitly copied summary."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("diagnosticView")
        self.details = QPlainTextEdit(self)
        self.details.setReadOnly(True)
        self.details.setAccessibleName(t("diagnostics.local_details"))
        self.addTab(self.details, t("diagnostics.local_details"))
        preview_page = QWidget(self)
        layout = QVBoxLayout(preview_page)
        layout.setContentsMargins(0, Spacing.CONTROL_GAP, 0, 0)
        note = QLabel(t("diagnostics.redacted_hint"), preview_page)
        note.setWordWrap(True)
        note.setTextFormat(Qt.TextFormat.PlainText)
        note.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        layout.addWidget(note)
        self.preview = QPlainTextEdit(preview_page)
        self.preview.setObjectName("diagnosticPreview")
        self.preview.setReadOnly(True)
        self.preview.setAccessibleName(t("diagnostics.preview"))
        layout.addWidget(self.preview, 1)
        self.addTab(preview_page, t("diagnostics.preview"))

    def set_content(self, details: str, diagnostic: DiagnosticSummary | None) -> None:
        for widget, text in ((self.details, details), (self.preview, diagnostic.to_text() if diagnostic else "")):
            if widget.toPlainText() != text:
                widget.setPlainText(text)

    def copy_diagnostics(self) -> None:
        text = self.preview.toPlainText()
        if text:
            QApplication.clipboard().setText(text)


class DiagnosticDialog(QDialog):
    """A resizable preview; copying never dismisses a recovery decision."""

    def __init__(
        self,
        title: str,
        message: str,
        *,
        details: str,
        diagnostic: DiagnosticSummary,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("feedbackDiagnosticDialog")
        self.setWindowTitle(title)
        self.resize(620, 440)
        self.setMinimumSize(360, 300)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(
            Spacing.CARD_PADDING, Spacing.CARD_PADDING, Spacing.CARD_PADDING, Spacing.CARD_PADDING
        )
        layout.setSpacing(Spacing.GROUP_GAP)
        message_label = QLabel(message, self)
        message_label.setWordWrap(True)
        message_label.setTextFormat(Qt.TextFormat.PlainText)
        message_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        layout.addWidget(message_label)
        self.view = DiagnosticView(self)
        self.view.set_content(details, diagnostic)
        self.view.setCurrentIndex(1)
        layout.addWidget(self.view, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close, self)
        close = buttons.button(QDialogButtonBox.StandardButton.Close)
        close.setText(t("common.close"))
        self.copy = buttons.addButton(t("diagnostics.copy"), QDialogButtonBox.ButtonRole.ActionRole)
        for button in buttons.buttons():
            button.setObjectName("secondaryActionButton")
            apply_theme_class(button, "secondary")
        self.copy.clicked.connect(self.view.copy_diagnostics)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        QWidget.setTabOrder(self.view.preview, self.copy)
        QWidget.setTabOrder(self.copy, close)
