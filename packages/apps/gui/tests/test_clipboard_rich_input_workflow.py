"""Rich clipboard GUI behavior split from the base clipboard lifecycle suite."""

from __future__ import annotations

from pathlib import Path

import pytest
from PySide6.QtCore import QMimeData, Qt
from PySide6.QtWidgets import QApplication

from docwen_core.models.clipboard_document import (
    ClipboardParagraph,
    ClipboardTable,
    clipboard_cell_text,
    clipboard_paragraph_text,
    load_clipboard_document_bytes,
)
from docwen_gui.main_window import MainWindow
from docwen_gui.view_models.main_window_vm import MainWindowViewModel

pytestmark = pytest.mark.gui


@pytest.fixture
def clipboard_window(qapp: QApplication, qtbot, tmp_path: Path):
    window = MainWindow(
        view_model=MainWindowViewModel(controller=None),
        clipboard_input_root=tmp_path / "clipboard-inputs",
    )
    qtbot.addWidget(window)
    window.resize(720, 620)
    window.show()
    qapp.processEvents()
    yield window
    window.close()
    qapp.clipboard().clear()
    qapp.processEvents()


def _paste_text(window: MainWindow, qapp: QApplication, qtbot, text: str) -> str:
    qapp.clipboard().setText(text)
    qtbot.mouseClick(window.input_area.paste_button, Qt.MouseButton.LeftButton)
    qtbot.waitUntil(lambda: not window.view_model.inspection_busy)
    qtbot.waitUntil(lambda: bool(window.view_model.files))
    selected = window.view_model.selected_file
    assert selected is not None
    return selected.path


def test_rich_clipboard_preserves_multiple_tables_and_body_order(
    clipboard_window: MainWindow,
    qapp: QApplication,
    qtbot,
) -> None:
    mime = QMimeData()
    mime.setText("  Before https://example.test/a **plain-md**  \nA\tB\n\t00123\nMiddle\nPipe\tLines\nx|y\tone\ntwo\nAfter  ")
    mime.setHtml(
        "<p>Before</p><table><tr><th>A</th><th>B</th></tr><tr><td></td><td>00123</td></tr></table>"
        "<p>Middle</p><table><tr><th>Pipe</th><th>Lines</th></tr>"
        "<tr><td>x|y</td><td>one<br>two</td></tr></table><p>After</p>"
    )
    qapp.clipboard().setMimeData(mime)
    qtbot.mouseClick(clipboard_window.input_area.paste_button, Qt.MouseButton.LeftButton)
    qtbot.waitUntil(lambda: not clipboard_window.view_model.inspection_busy)
    selected = clipboard_window.view_model.selected_file
    assert selected is not None
    document = load_clipboard_document_bytes(Path(selected.path).read_bytes())
    assert [isinstance(block, ClipboardTable) for block in document.blocks] == [False, True, False, True, False]
    assert [clipboard_paragraph_text(block) for block in document.blocks if isinstance(block, ClipboardParagraph)] == [
        "  Before https://example.test/a **plain-md**  \n",
        "\nMiddle\n",
        "\nAfter  ",
    ]
    tables = [block for block in document.blocks if isinstance(block, ClipboardTable)]
    assert [clipboard_cell_text(cell) for cell in tables[0].cells] == ["A", "B", "", "00123"]
    assert [clipboard_cell_text(cell) for cell in tables[1].cells] == ["Pipe", "Lines", "x|y", "one\ntwo"]
    assert not any(row.message_type in {"warning", "danger"} for row in clipboard_window._info_area_vm.history_rows)


def test_valid_rowspan_is_preserved_without_fallback_warning(
    clipboard_window: MainWindow,
    qapp: QApplication,
    qtbot,
) -> None:
    mime = QMimeData()
    mime.setText("A\tB\nkept-rowspan\tfirst\n\tsecond")
    mime.setHtml(
        "<table><tr><th>A</th><th>B</th></tr>"
        "<tr><td rowspan='2'>kept-rowspan</td><td>first</td></tr><tr><td>second</td></tr></table>"
    )
    qapp.clipboard().setMimeData(mime)
    qtbot.mouseClick(clipboard_window.input_area.paste_button, Qt.MouseButton.LeftButton)
    qtbot.waitUntil(lambda: not clipboard_window.view_model.inspection_busy)
    selected = clipboard_window.view_model.selected_file
    assert selected is not None
    document = load_clipboard_document_bytes(Path(selected.path).read_bytes())
    table = document.blocks[0]
    assert isinstance(table, ClipboardTable)
    assert table.row_count == 3 and table.column_count == 2
    assert [clipboard_cell_text(cell) for cell in table.cells] == ["A", "B", "kept-rowspan", "first", "second"]
    assert next(cell for cell in table.cells if clipboard_cell_text(cell) == "kept-rowspan").row_span == 2
    assert not any(row.message_type in {"warning", "danger"} for row in clipboard_window._info_area_vm.history_rows)


def test_invalid_rowspan_rejects_paste_and_preserves_existing_input(
    clipboard_window: MainWindow,
    qapp: QApplication,
    qtbot,
) -> None:
    original = _paste_text(clipboard_window, qapp, qtbot, "original text")
    mime = QMimeData()
    mime.setText("malformed table text")
    mime.setHtml("<table><tr><td rowspan='2'>outside declared rows</td></tr></table>")
    qapp.clipboard().setMimeData(mime)
    qtbot.mouseClick(clipboard_window.input_area.paste_button, Qt.MouseButton.LeftButton)
    qapp.processEvents()
    selected = clipboard_window.view_model.selected_file
    assert selected is not None and selected.path == original
    assert Path(original).read_text(encoding="utf-8") == "original text"
    assert any(row.message_type == "danger" for row in clipboard_window._info_area_vm.history_rows)


def test_plain_markdown_paste_remains_exact_text(
    clipboard_window: MainWindow,
    qapp: QApplication,
    qtbot,
) -> None:
    raw = "# Heading\n\n| A | B |\n| --- | --- |\n| x\\|y | 00123 |\n"
    selected = _paste_text(clipboard_window, qapp, qtbot, raw)
    assert Path(selected).read_text(encoding="utf-8") == raw


def test_visible_plain_text_menu_bypasses_rich_projection(
    clipboard_window: MainWindow,
    qapp: QApplication,
    qtbot,
) -> None:
    plain = "Name\tValue\nA\t00123"
    mime = QMimeData()
    mime.setText(plain)
    mime.setHtml("<table><tr><th>Name</th><th>Value</th></tr><tr><td>A</td><td>00123</td></tr></table>")
    qapp.clipboard().setMimeData(mime)
    button = clipboard_window.input_area.paste_menu_button
    assert button.isVisible()
    menu = button.menu()
    assert menu is not None
    actions = menu.actions()
    assert len(actions) == 1 and actions[0].text()
    actions[0].trigger()
    qtbot.waitUntil(lambda: not clipboard_window.view_model.inspection_busy)
    selected = clipboard_window.view_model.selected_file
    assert selected is not None
    assert Path(selected.path).read_text(encoding="utf-8") == plain


def test_rich_clipboard_reports_images_without_importing_them(
    clipboard_window: MainWindow,
    qapp: QApplication,
    qtbot,
) -> None:
    mime = QMimeData()
    mime.setText("Before\nAfter")
    mime.setHtml("<p>Before</p><img src='https://example.invalid/private.png'><p>After</p>")
    qapp.clipboard().setMimeData(mime)
    qtbot.mouseClick(clipboard_window.input_area.paste_button, Qt.MouseButton.LeftButton)
    qtbot.waitUntil(lambda: not clipboard_window.view_model.inspection_busy)
    selected = clipboard_window.view_model.selected_file
    assert selected is not None
    assert Path(selected.path).read_text(encoding="utf-8") == "Before\nAfter"
    assert any(row.message_type == "warning" for row in clipboard_window._info_area_vm.history_rows)



def test_reliable_plain_alignment_preserves_nested_merged_table_structure_and_plain_boundaries(
    clipboard_window: MainWindow,
    qapp: QApplication,
    qtbot,
) -> None:
    mime = QMimeData()
    plain = "  Lead  \nA\tB\nbefore\ninner\nafter\t\nTail  "
    mime.setText(plain)
    mime.setHtml(
        "<p>HTML lead must not win</p><table><tr><th>A</th><th>B</th></tr>"
        "<tr><td colspan='2'><p>before</p><table><tr><td>inner</td></tr></table>"
        "<p>after</p></td></tr></table><p>HTML tail must not win</p>"
    )
    qapp.clipboard().setMimeData(mime)

    qtbot.mouseClick(clipboard_window.input_area.paste_button, Qt.MouseButton.LeftButton)
    qtbot.waitUntil(lambda: not clipboard_window.view_model.inspection_busy)
    selected = clipboard_window.view_model.selected_file
    assert selected is not None and selected.format == "clipboard_document"
    document = load_clipboard_document_bytes(Path(selected.path).read_bytes())
    assert isinstance(document.blocks[0], ClipboardParagraph)
    assert clipboard_paragraph_text(document.blocks[0]) == "  Lead  \n"
    table = document.blocks[1]
    assert isinstance(table, ClipboardTable)
    merged = next(cell for cell in table.cells if cell.row == 1 and cell.column == 0)
    assert merged.column_span == 2
    assert [type(block).__name__ for block in merged.blocks] == [
        "ClipboardParagraph",
        "ClipboardTable",
        "ClipboardParagraph",
    ]
    assert clipboard_paragraph_text(document.blocks[2]) == "\nTail  "


def test_unreliable_html_table_alignment_keeps_complete_plain_text(
    clipboard_window: MainWindow,
    qapp: QApplication,
    qtbot,
) -> None:
    plain = "  KEEP https://example.test/x **markdown**  \nnot the HTML table\nTAIL\u00a0 "
    mime = QMimeData()
    mime.setText(plain)
    mime.setHtml(
        "<p>different HTML body</p><table><tr><th>A</th><th>B</th></tr>"
        "<tr><td>1</td><td>2</td></tr></table><p>different tail</p>"
    )
    qapp.clipboard().setMimeData(mime)

    qtbot.mouseClick(clipboard_window.input_area.paste_button, Qt.MouseButton.LeftButton)
    qtbot.waitUntil(lambda: not clipboard_window.view_model.inspection_busy)
    selected = clipboard_window.view_model.selected_file
    assert selected is not None
    assert selected.format == "markdown"
    assert Path(selected.path).read_text(encoding="utf-8") == plain
    assert any(row.message_type == "warning" for row in clipboard_window._info_area_vm.history_rows)


def test_html_only_structured_table_is_explicitly_warned(
    clipboard_window: MainWindow,
    qapp: QApplication,
    qtbot,
) -> None:
    mime = QMimeData()
    mime.setHtml("<table><tr><th>A</th><th>B</th></tr><tr><td>1</td><td>2</td></tr></table>")
    qapp.clipboard().setMimeData(mime)

    qtbot.mouseClick(clipboard_window.input_area.paste_button, Qt.MouseButton.LeftButton)
    qtbot.waitUntil(lambda: not clipboard_window.view_model.inspection_busy)
    selected = clipboard_window.view_model.selected_file
    assert selected is not None and selected.format == "clipboard_document"
    assert any(row.message_type == "warning" for row in clipboard_window._info_area_vm.history_rows)
