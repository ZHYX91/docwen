"""Optional Mermaid installation controls and cancellable test rendering."""

from __future__ import annotations

import tempfile
from pathlib import Path
from typing import TYPE_CHECKING, Any

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices, QPixmap
from PySide6.QtWidgets import QFileDialog, QFormLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QWidget

from docwen_core.mermaid_render import render_mermaid_png
from docwen_core.mermaid_runtime import inspect_mermaid_runtime

from ...i18n import t
from ...qt_bridge.background_operation import BackgroundOperation
from ...view_models.settings_vm import SECTION_FORMATTING, SettingsViewModel

if TYPE_CHECKING:
    from .base_tab import BaseSettingsTab


class MermaidControls(QWidget):
    def __init__(self, owner: BaseSettingsTab, form: QFormLayout, model: SettingsViewModel) -> None:
        super().__init__(owner)
        self._model = model
        self.operation = BackgroundOperation(self)
        token_owner = self.operation
        self.destroyed.connect(lambda: token_owner.cancel())
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.path = QLineEdit()
        self.path.setObjectName("formattingMermaidPath")
        self.path.setPlaceholderText(t("settings.formatting.mermaid_auto"))
        browse = QPushButton(t("settings.formatting.mermaid_browse"))
        browse.clicked.connect(self._browse)
        layout.addWidget(self.path, 1)
        layout.addWidget(browse)
        owner.add_form_row(form, t("settings.formatting.mermaid_path"), self)
        self.status = QLabel()
        self.status.setWordWrap(True)
        self.status.setTextFormat(Qt.TextFormat.PlainText)
        self.status.setObjectName("formattingMermaidAvailability")
        form.addRow(self.status)
        for key, label, callback in (
            ("mermaid_recheck", t("settings.formatting.mermaid_recheck"), self.refresh),
            ("mermaid_test", t("settings.formatting.mermaid_test"), self.test_render),
            ("mermaid_install", t("settings.formatting.mermaid_install"), self._instructions),
        ):
            button = QPushButton(label)
            button.setObjectName(key)
            button.clicked.connect(callback)
            form.addRow(button)
        self.preview = QLabel()
        self.preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview.setObjectName("formattingMermaidPreview")
        form.addRow(self.preview)
        self.path.textChanged.connect(self._path_changed)

    def load(self, path: str) -> None:
        self.path.blockSignals(True)
        self.path.setText(path)
        self.path.blockSignals(False)
        self.refresh()

    def _path_changed(self, value: str) -> None:
        self._model.set_field(SECTION_FORMATTING, "mermaid_cli_path", value.strip())
        self.refresh()

    def _browse(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, t("settings.formatting.mermaid_browse"), self.path.text())
        if path:
            self.path.setText(path)

    @staticmethod
    def _instructions() -> None:
        QDesktopServices.openUrl(QUrl("https://github.com/ZHYX91/docwen/blob/main/docs/mermaid.md"))

    def refresh(self) -> None:
        self.operation.cancel()
        self.preview.clear()
        runtime = inspect_mermaid_runtime(self.path.text())
        if runtime.available:
            message = t("settings.formatting.mermaid_ready", cli=runtime.cli_version, mermaid=runtime.mermaid_version)
        elif runtime.reason == "unsupported_version":
            message = t(
                "settings.formatting.mermaid_incompatible", cli=runtime.cli_version, mermaid=runtime.mermaid_version
            )
        else:
            message = t("settings.formatting.mermaid_unavailable")
        self.status.setText(message)

    def test_render(self) -> None:
        path = self.path.text()
        # The caller owns the selected process scratch root. Governed QA and
        # source acceptance bind TEMP to their leased run before launching Qt.
        root = Path(tempfile.gettempdir())
        self.preview.clear()
        self.status.setText(t("settings.formatting.mermaid_testing"))
        self.operation.submit(
            lambda cancellation: render_mermaid_png(
                "flowchart LR\nA[Mermaid] --> B[OK]", work_dir=root, cli_path=path, cancellation=cancellation
            ),
            self._completed,
        )

    def _completed(self, data: Any, error: Exception | None) -> None:
        if error is not None:
            message = (
                t("settings.formatting.mermaid_browser_missing")
                if "could not find chrome" in str(error).lower()
                else t("settings.formatting.mermaid_test_failed", reason=str(error)[:1200])
            )
            self.status.setText(message)
            return
        pixmap = QPixmap()
        if not pixmap.loadFromData(data):
            self.status.setText(t("settings.formatting.mermaid_test_failed", reason="PNG"))
            return
        self.preview.setPixmap(
            pixmap.scaled(260, 90, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
        )
        self.status.setText(t("settings.formatting.mermaid_test_passed"))

    def hideEvent(self, event) -> None:
        self.operation.cancel()
        super().hideEvent(event)
