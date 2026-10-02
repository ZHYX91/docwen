"""Actual paste and batch updates must reach the visible generation picker."""

from pathlib import Path

import pytest
from PySide6.QtCore import QMimeData, Qt

from docwen_application.controller import ApplicationController
from docwen_bundle.runtime_factory import create_runtime_port
from docwen_gui.main_window import MainWindow
from docwen_gui.view_models.main_window_vm import MainWindowViewModel

pytestmark = pytest.mark.gui


def test_rich_paste_and_mixed_batch_refresh_visible_targets(qapp, qtbot, tmp_path: Path, monkeypatch):
    monkeypatch.setattr("docwen_bundle.runtime_factory._runtime_workspace_root", lambda: tmp_path / "runtime")
    controller = ApplicationController(runtime_port=create_runtime_port())
    controller.start()
    window = MainWindow(
        view_model=MainWindowViewModel(controller=controller), clipboard_input_root=tmp_path / "clipboard"
    )
    qtbot.addWidget(window)
    window.show()
    try:
        mime = QMimeData()
        mime.setText("A\tB\n00123\t2")
        mime.setHtml("<table><tr><td>A</td><td>B</td></tr><tr><td>00123</td><td>2</td></tr></table>")
        qapp.clipboard().setMimeData(mime)
        qtbot.mouseClick(window.input_area.paste_button, Qt.MouseButton.LeftButton)
        qtbot.waitUntil(lambda: not window.view_model.inspection_busy and window.view_model.selected_file is not None)
        selected = window.view_model.selected_file
        assert selected is not None and selected.format == "clipboard_document"
        combo = window._action_area.md_document_format_combo
        assert combo.findData("md") >= 0
        combo.setCurrentIndex(combo.findData("md"))
        assert window._action_area_vm.target_format == "md"
        window._action_area_vm.set_template_ready(False)
        assert window._action_area.convert_docx_button.isEnabled()
        assert window._action_area.md_remove_numbering_cb is None
        assert window._action_area_vm.collect_options() == {}

        ordinary = tmp_path / "ordinary.md"
        ordinary.write_text("# Ordinary Markdown\n", encoding="utf-8")
        window.view_model.mode = "batch"
        window.view_model.add_files([str(ordinary)])
        window.view_model.set_selected_file(selected)
        window._show_template_target_mode("docx", selected.path)
        combo = window._action_area.md_document_format_combo
        assert combo.findData("md") == -1
        assert combo.findData("docx") >= 0
    finally:
        window.view_model.cancel_inspection()
        window.close()
        qapp.clipboard().clear()
        controller.stop()
