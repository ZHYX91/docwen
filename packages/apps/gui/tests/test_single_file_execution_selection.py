"""Keep the visible input, canonical selection and conversion request aligned."""

from pathlib import Path

import pytest

from ._main_window_projection_binding_support import _load_request_templates
from ._main_window_projection_binding_support import window as window

pytestmark = pytest.mark.gui


def _sources(tmp_path):
    paths = [tmp_path / "first.md", tmp_path / "second.md"]
    for path in paths:
        path.write_text(f"# {path.stem}\n", encoding="utf-8")
    return paths


@pytest.mark.parametrize("reselect_existing", [False, True])
def test_single_add_executes_the_visible_file(window, tmp_path, monkeypatch, reselect_existing):
    first, second = _sources(tmp_path)
    calls = []
    monkeypatch.setattr(window, "_start_execution", lambda **kwargs: calls.append(kwargs))
    _load_request_templates(window)
    window._input_area_vm.add_files([str(first)])
    window._input_area_vm.add_files([str(second)])
    expected = second
    if reselect_existing:
        window._input_area_vm.add_files([str(first)])
        expected = first

    assert len(window._view_model.files) == 2
    assert Path(window._view_model.selected_file.path) == expected
    assert expected.name in window._input_area_vm.selection_message
    assert Path(window._batch_list.get_current_file()) == expected
    window._action_area_vm.request_conversion()
    assert len(calls) == 1
    assert Path(calls[0]["file_path"]) == expected


def test_rejected_single_add_preserves_execution_selection(window, tmp_path, monkeypatch):
    first, _second = _sources(tmp_path)
    _load_request_templates(window)
    window._input_area_vm.add_files([str(first)])
    calls = []
    monkeypatch.setattr(window, "_start_execution", lambda **kwargs: calls.append(kwargs))
    window._input_area_vm.add_files([str(tmp_path / "missing.md")])
    assert Path(window._view_model.selected_file.path) == first
    window._action_area_vm.request_conversion()
    assert Path(calls[0]["file_path"]) == first


def test_batch_add_preserves_selection_and_single_mode_projects_it(window, tmp_path):
    first, second = _sources(tmp_path)
    window._input_area_vm.add_files([str(first)])
    window._input_area_vm.set_mode("batch")
    window._input_area_vm.add_files([str(second)])
    assert Path(window._view_model.selected_file.path) == first
    window._batch_list.select_file(str(second))
    window._input_area_vm.set_mode("single")
    assert Path(window._view_model.selected_file.path) == second
    assert second.name in window._input_area_vm.selection_message


def test_batch_projection_preserves_multiple_selected_rows(window, tmp_path):
    first, second = _sources(tmp_path)
    window._input_area_vm.set_mode("batch")
    window._input_area_vm.add_files([str(first), str(second)])
    window._batch_list._tabs["text"].selectAll()
    window._view_model.ui_projection_changed.emit(window._view_model.ui_projection)
    assert {Path(path) for path in window._batch_list.get_selected_files("text")} == {first, second}
