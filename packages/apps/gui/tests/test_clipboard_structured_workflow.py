"""GUI user path for structured table clipboard capture."""

from __future__ import annotations

from pathlib import Path

import pytest
from PySide6.QtCore import QMimeData, Qt
from PySide6.QtWidgets import QApplication

from docwen_core.models.clipboard_document import load_clipboard_document_bytes
from docwen_gui.main_window import MainWindow
from docwen_gui.view_models.main_window_vm import MainWindowViewModel

pytestmark = pytest.mark.gui


@pytest.fixture
def window(qapp: QApplication, qtbot, tmp_path: Path):
    value = MainWindow(
        view_model=MainWindowViewModel(controller=None),
        clipboard_input_root=tmp_path / "clipboard-inputs",
    )
    qtbot.addWidget(value)
    value.resize(720, 620)
    value.show()
    qapp.processEvents()
    yield value
    value.close()
    qapp.clipboard().clear()
    qapp.processEvents()


def test_default_table_paste_creates_one_structured_managed_input(window, qapp, qtbot) -> None:
    mime = QMimeData()
    mime.setText("Before\nA\tB\n1\tbefore\nnested\nafter\n\t2\nAfter")
    mime.setHtml(
        "<p>Before</p><table><tr><th>A</th><th>B</th></tr>"
        "<tr><td rowspan='2'>1</td><td><p>before</p>"
        "<table><tr><td>nested</td></tr></table><p>after</p></td></tr>"
        "<tr><td>2</td></tr></table><p>After</p>"
    )
    qapp.clipboard().setMimeData(mime)

    qtbot.mouseClick(window.input_area.paste_button, Qt.MouseButton.LeftButton)
    qtbot.waitUntil(lambda: not window.view_model.inspection_busy)
    qtbot.waitUntil(lambda: len(window.view_model.files) == 1)

    ref = window.view_model.files[0]
    assert ref.format == "clipboard_document"
    assert ref.category == "markdown"
    assert Path(ref.path).suffix == ".dwclip"
    assert window._clipboard_store is not None
    bundle = window._clipboard_store.bundle(ref.path)
    assert bundle is not None
    assert bundle.resources == ()
    document = load_clipboard_document_bytes(Path(ref.path).read_bytes())
    assert len(document.blocks) == 3
    assert window._clipboard_store.snapshot_available(ref.path)


def test_invalid_table_structure_keeps_complete_plain_text_with_warning(window, qapp, qtbot) -> None:
    mime = QMimeData()
    plain = "  A B C\nhttps://example.test/a **authored**\u00a0 "
    mime.setText(plain)
    mime.setHtml("<table><tr><td>A</td><td>B</td></tr><tr><td>C</td></tr></table>")
    qapp.clipboard().setMimeData(mime)

    qtbot.mouseClick(window.input_area.paste_button, Qt.MouseButton.LeftButton)
    qtbot.waitUntil(lambda: not window.view_model.inspection_busy)
    assert len(window.view_model.files) == 1
    ref = window.view_model.files[0]
    assert ref.format == "markdown"
    assert Path(ref.path).read_text(encoding="utf-8") == plain
    assert any(row.message_type == "warning" for row in window._info_area_vm.history_rows)
    assert not any(row.message_type == "danger" for row in window._info_area_vm.history_rows)


def test_invalid_table_structure_without_plain_body_is_rejected(window, qapp, qtbot) -> None:
    mime = QMimeData()
    mime.setHtml("<table><tr><td>A</td><td>B</td></tr><tr><td>C</td></tr></table>")
    qapp.clipboard().setMimeData(mime)

    qtbot.mouseClick(window.input_area.paste_button, Qt.MouseButton.LeftButton)
    qapp.processEvents()

    assert window.view_model.files == []
    assert any(row.message_type == "danger" for row in window._info_area_vm.history_rows)


def test_plain_text_action_does_not_create_structured_input(window, qapp, qtbot) -> None:
    mime = QMimeData()
    plain = "A\tB\n1\t2"
    mime.setText(plain)
    mime.setHtml("<table><tr><td>A</td><td>B</td></tr><tr><td>1</td><td>2</td></tr></table>")
    qapp.clipboard().setMimeData(mime)

    window.input_area.request_plain_text_paste()
    qtbot.waitUntil(lambda: not window.view_model.inspection_busy)
    ref = window.view_model.files[0]
    assert ref.format == "markdown"
    assert Path(ref.path).read_text(encoding="utf-8") == plain
