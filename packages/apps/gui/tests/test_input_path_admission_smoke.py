"""Cross-platform release smoke for real file identity at the GUI admission boundary."""

from pathlib import Path

import pytest
from PySide6.QtCore import QMimeData, Qt, QUrl

from docwen_gui.main_window import MainWindow
from docwen_gui.view_models.main_window_vm import MainWindowViewModel

pytestmark = [pytest.mark.gui_smoke, pytest.mark.release_gate]


@pytest.mark.parametrize("names", [("A.md", "a.md"), ("Straße.md", "Strasse.md")])
@pytest.mark.parametrize("source", ["files", "folder", "text", "clipboard"])
def test_batch_preserves_distinct_casefold_colliding_files(qapp, qtbot, tmp_path, names, source) -> None:
    inputs = tmp_path / "inputs"
    inputs.mkdir()
    first, second = (inputs / name for name in names)
    first.write_text("# First\n", encoding="utf-8")
    second.write_text("# Second\n", encoding="utf-8")
    if first.samefile(second):
        pytest.skip("The actual filesystem does not distinguish these filenames")
    window = MainWindow(
        view_model=MainWindowViewModel(controller=None),
        clipboard_input_root=tmp_path / "clipboard-inputs",
    )
    qtbot.addWidget(window)
    window.show()
    vm = window._input_area_vm
    vm.set_mode("batch")
    paths = [str(first), str(second), str(first)]
    if source == "folder":
        paths = [str(inputs), str(first)]
    elif source == "text":
        paths = vm.extract_paths_from_text_payload("\n".join(paths))
        assert paths == [str(first), str(second)]
    preview = vm.build_drag_preview(paths)
    assert preview.added_count == 2
    assert preview.skipped_count == 0
    if source == "clipboard":
        mime = QMimeData()
        mime.setUrls([QUrl.fromLocalFile(path) for path in paths])
        qapp.clipboard().setMimeData(mime)
        qtbot.mouseClick(window.input_area.paste_button, Qt.MouseButton.LeftButton)
    else:
        vm.add_files(paths)
    qtbot.waitUntil(lambda: not window.view_model.inspection_busy)
    qtbot.waitUntil(lambda: len(window.view_model.files) == 2)
    assert {Path(ref.path).name for ref in window.view_model.files} == set(names)
    assert window._clipboard_store is None
    assert first.read_text(encoding="utf-8") == "# First\n"
    assert second.read_text(encoding="utf-8") == "# Second\n"
    window.close()
