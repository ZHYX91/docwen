"""GUI 逻辑单元测试。"""

from __future__ import annotations

import pytest

from docwen.gui.core.event_handler import MainWindowEventHandler

pytestmark = [pytest.mark.unit, pytest.mark.windows_only]


class _FileDropArea:
    def __init__(self, mode: str):
        self._mode = mode

    def get_mode(self) -> str:
        return self._mode


class _MainWindow:
    def __init__(self, mode: str):
        self.file_drop_area = _FileDropArea(mode)

    def _show_error(self, *_args, **_kwargs) -> None:
        raise AssertionError("unexpected error")


class _Logic:
    def __init__(self):
        self.calls = []

    def handle_batch_files_added(self, files):
        self.calls.append(("batch", list(files)))

    def handle_file_dropped(self, file_path, *, mode: str):
        self.calls.append(("single", file_path, mode))


def test_event_handler_on_file_dropped_list_single_mode() -> None:
    handler = MainWindowEventHandler.__new__(MainWindowEventHandler)
    handler.main_window = _MainWindow("single")
    handler.logic = _Logic()

    handler.on_file_dropped(["a.docx"])

    assert handler.logic.calls == [("single", "a.docx", "single")]


def test_event_handler_on_file_dropped_list_batch_mode() -> None:
    handler = MainWindowEventHandler.__new__(MainWindowEventHandler)
    handler.main_window = _MainWindow("batch")
    handler.logic = _Logic()

    handler.on_file_dropped(["a.docx", "b.docx"])

    assert handler.logic.calls == [("batch", ["a.docx", "b.docx"])]


def test_file_drop_area_simulate_drop_single_mode_passes_list() -> None:
    pytest.importorskip("ttkbootstrap")
    from docwen.gui.components.file_drop import FileDropArea

    captured = []

    class _Dummy:
        mode = "single"
        file_path = None

        def _validate_single_file_mode(self, file_paths):
            return True, ""

        def _switch_to_file_state(self):
            return None

    dummy = _Dummy()
    dummy.on_file_dropped = lambda files: captured.append(files)

    FileDropArea._simulate_drop(dummy, ["a.docx"])

    assert captured == [["a.docx"]]
