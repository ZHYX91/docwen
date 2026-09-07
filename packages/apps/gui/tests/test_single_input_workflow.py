"""One visible input across drag/picker, IPC, and mode switches."""

from pathlib import Path

import pytest

from docwen_gui.i18n import t

pytestmark = pytest.mark.gui


def _files(tmp_path):
    paths = [tmp_path / name for name in ("first.md", "second.md", "third.md")]
    for path in paths:
        path.write_text(f"# {path.stem}\n\nKeep this original.\n", encoding="utf-8")
    return paths


@pytest.mark.parametrize("entry", ["input", "ipc", "model"])
def test_single_input_replaces_previous_and_never_resurrects_it(main_window, tmp_path, entry):
    window = main_window
    vm = window._view_model
    paths = _files(tmp_path)
    history = list(window._info_area_vm.history_rows)
    for path in (*paths, paths[0]):
        if entry == "input":
            window._input_area_vm.add_files([str(path)])
        elif entry == "ipc":
            assert window.handle_ipc_command("open_file", str(path))
        else:
            vm.add_files([str(path)])
        assert [Path(ref.path) for ref in vm.files] == [path]
        assert Path(vm.selected_file.path) == path
        assert [Path(p) for p in window._batch_list_vm.get_files()] == [path]
        assert path.name in window._input_area_vm.selection_message
        assert Path(window._action_area_vm._file_path) == path
    assert window._info_area_vm.history_rows == history
    window._input_area_vm.set_mode("batch")
    assert [Path(ref.path) for ref in vm.files] == [paths[0]]
    assert all("Keep this original." in path.read_text(encoding="utf-8") for path in paths)


@pytest.mark.parametrize("accept", [False, True])
def test_batch_to_single_confirmation_preserves_or_reduces_visible_input(main_window, tmp_path, monkeypatch, accept):
    window = main_window
    paths = _files(tmp_path)
    window._input_area_vm.set_mode("batch")
    window._input_area_vm.add_files([str(path) for path in paths])
    window._batch_list.select_file(str(paths[1]))
    calls = []

    def confirm(title, message, **kwargs):
        calls.append((title, message))
        return accept

    monkeypatch.setattr("docwen_gui.dialogs.feedback.confirm", confirm)
    window._input_area._request_mode("single")
    assert len(calls) == 1
    assert paths[1].name in calls[0][1]
    assert window._view_model.mode == ("single" if accept else "batch")
    assert window._input_area_vm.mode == window._view_model.mode
    assert [Path(ref.path) for ref in window._view_model.files] == ([paths[1]] if accept else paths)
    assert Path(window._view_model.selected_file.path) == paths[1]
    window._input_area_vm.set_mode("batch")
    assert [Path(p) for p in window._batch_list_vm.get_files()] == ([paths[1]] if accept else paths)
    assert all(path.exists() for path in paths)


def test_rejected_replacement_and_multiple_inputs_preserve_current_file(main_window, tmp_path):
    vm = main_window._view_model
    paths = _files(tmp_path)
    vm.add_files([str(paths[0])])
    for replacement in ([str(tmp_path / "missing.md")], [str(paths[1]), str(paths[2])]):
        outcome = vm.add_files(replacement)
        assert outcome.rejected and not outcome.added
        assert [Path(ref.path) for ref in vm.files] == [paths[0]]
        assert Path(vm.selected_file.path) == paths[0]
        assert paths[0].name in main_window._input_area_vm.selection_message
        main_window._input_area_vm.add_files(replacement)
        assert paths[0].name in main_window._input_area_vm.selection_message


def test_batch_ipc_adds_to_visible_list_and_selects_received_file(main_window, tmp_path):
    vm = main_window._view_model
    paths = _files(tmp_path)
    main_window._input_area_vm.set_mode("batch")
    for path in paths:
        assert main_window.handle_ipc_command("open_file", str(path))
        assert Path(vm.selected_file.path) == path
    assert [Path(ref.path) for ref in vm.files] == paths
    assert [Path(p) for p in main_window._batch_list_vm.get_files()] == paths


def test_running_task_refuses_input_and_mode_changes_until_owner_releases(main_window, tmp_path):
    vm = main_window._view_model
    paths = _files(tmp_path)
    vm.add_files([str(paths[0])])
    vm.reserve_execution_inputs("active", (str(paths[0]),))
    assert not main_window.handle_ipc_command("open_file", str(paths[1]))
    main_window._input_area_vm.set_mode("batch")
    assert vm.mode == main_window._input_area_vm.mode == "single"
    assert [Path(ref.path) for ref in vm.files] == [paths[0]]
    assert Path(vm.selected_file.path) == paths[0]
    assert vm.status_message == t("components.file_drop.input_busy")
    vm.release_execution_inputs("active")
    assert main_window.handle_ipc_command("open_file", str(paths[1]))
    assert [Path(ref.path) for ref in vm.files] == [paths[1]]
