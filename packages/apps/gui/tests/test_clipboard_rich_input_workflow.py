"""Rich clipboard GUI behavior split from the base clipboard lifecycle suite."""

from __future__ import annotations

from pathlib import Path

import pytest
from PySide6.QtCore import QMimeData, Qt
from PySide6.QtWidgets import QApplication

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
    mime.setText("Before\nA\tB\n\t00123\nMiddle\nPipe\tLines\nx|y\tone two\nAfter")
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
    content = Path(selected.path).read_text(encoding="utf-8")
    assert content.index("Before") < content.index("| A | B |") < content.index("Middle")
    assert content.index("Middle") < content.index("| Pipe | Lines |") < content.index("After")
    assert "|  | 00123 |" in content
    assert "| x\\|y | one<br>two |" in content


def test_unsafe_table_keeps_extractable_text_and_warns(
    clipboard_window: MainWindow,
    qapp: QApplication,
    qtbot,
) -> None:
    mime = QMimeData()
    mime.setText("A\tB\nkept-rowspan\tfirst\nsecond")
    mime.setHtml(
        "<table><tr><th>A</th><th>B</th></tr>"
        "<tr><td rowspan='2'>kept-rowspan</td><td>first</td></tr><tr><td>second</td></tr></table>"
    )
    qapp.clipboard().setMimeData(mime)
    qtbot.mouseClick(clipboard_window.input_area.paste_button, Qt.MouseButton.LeftButton)
    qtbot.waitUntil(lambda: not clipboard_window.view_model.inspection_busy)
    selected = clipboard_window.view_model.selected_file
    assert selected is not None
    text = Path(selected.path).read_text(encoding="utf-8")
    assert all(value in text for value in ("A", "B", "kept-rowspan", "first", "second"))
    assert any(row.message_type == "warning" for row in clipboard_window._info_area_vm.history_rows)


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
