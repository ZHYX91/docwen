"""Native file-dialog results must reach shared single/batch admission intact."""

from pathlib import Path

import pytest
from PySide6.QtWidgets import QFileDialog

from docwen_gui.view_models.input_area_vm import InputAreaViewModel
from docwen_gui.view_models.main_window_vm import MainWindowViewModel
from docwen_gui.widgets.input_area import InputArea

pytestmark = pytest.mark.gui


@pytest.mark.parametrize("mode", ["single", "batch"])
def test_file_dialog_preserves_complete_selection_for_admission(qtbot, monkeypatch, tmp_path: Path, mode: str) -> None:
    paths = [tmp_path / name for name in ("original.md", "first.md", "second.md")]
    for path in paths:
        path.write_text("# Controlled input", encoding="utf-8")
    main = MainWindowViewModel(controller=None)
    main.add_files([str(paths[0])])
    vm = InputAreaViewModel(main_vm=main)
    vm.set_mode(mode)
    widget = InputArea(view_model=vm)
    qtbot.addWidget(widget)
    selected = [str(path) for path in paths[1:]]
    monkeypatch.setattr(QFileDialog, "getOpenFileNames", lambda *args: (selected, ""))
    monkeypatch.setattr(QFileDialog, "getOpenFileName", lambda *args: (selected[0], ""))
    admitted = []
    vm.files_added.connect(admitted.append)

    widget.open_file_dialog()
    qtbot.waitUntil(lambda: not main.inspection_busy)

    if mode == "single":
        assert admitted == []
        assert main.selected_file is not None
        assert main.selected_file.path == str(paths[0])
        assert vm.selection_tone == "warning"
    else:
        assert admitted == [selected]
        assert {ref.path for ref in main.files} == {str(path) for path in paths}
