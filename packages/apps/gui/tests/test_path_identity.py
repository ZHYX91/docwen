"""Native path spellings must keep distinct inputs distinct through the GUI."""

from __future__ import annotations

import ntpath
import os
import posixpath
from pathlib import PurePosixPath, PureWindowsPath
from types import SimpleNamespace

import pytest

from docwen_core.models.file_ref import FileRef
from docwen_gui import path_identity
from docwen_gui.execution_requests import ExecutionRequestBuilder
from docwen_gui.path_identity import normalize_path, path_key
from docwen_gui.view_models.batch_list_vm import BatchListViewModel
from docwen_gui.view_models.main_window_vm import MainWindowViewModel
from docwen_gui.view_models.task_history import TaskHistory
from docwen_gui.widgets.batch_list import BatchList

pytestmark = pytest.mark.unit


def _use_path_flavor(monkeypatch, path_module, separator):
    # Isolate only this module's host policy; never patch global os/pathlib.
    monkeypatch.setattr(path_identity, "os", SimpleNamespace(path=path_module, sep=separator, fspath=os.fspath))
    monkeypatch.setattr(path_identity, "Path", PurePosixPath if separator == "/" else PureWindowsPath)


def test_posix_backslash_is_part_of_filename(monkeypatch):
    _use_path_flavor(monkeypatch, posixpath, "/")
    assert normalize_path(r"/vault/a\b.md") == r"/vault/a\b.md"
    assert path_key(r"/vault/a\b.md") != path_key("/vault/a/b.md")
    assert path_key("/vault/a/../b.md") == path_key("/vault/b.md")
    assert path_key("/vault/A.md") != path_key("/vault/a.md")
    assert path_key("") == normalize_path("") == ""


def test_windows_alias_keys_preserve_source_spelling(monkeypatch):
    _use_path_flavor(monkeypatch, ntpath, "\\")
    assert path_key(r"C:\Vault\Note.md") == path_key("c:/vault/./NOTE.md")
    assert normalize_path(r"C:\Vault\Note.md") == "C:/Vault/Note.md"
    vm = BatchListViewModel()

    def resolver(_path):
        return {"detected_format": "markdown", "workflow_category": "markdown"}

    added, failed = vm.add_files([r"C:\Vault\Note.md", "c:/vault/NOTE.md"], file_resolver=resolver)
    assert added == ["C:/Vault/Note.md"]
    assert not failed
    entry = vm.get_file_entry("c:/vault/NOTE.md")
    assert entry is not None
    assert entry.file_path == "C:/Vault/Note.md"
    vm.set_file_status("c:/vault/NOTE.md", "failed")
    assert vm.get_failed_files() == ["C:/Vault/Note.md"]


def _assert_distinct_gui_inputs(window, paths):
    first, second = paths
    refs = [
        FileRef(path=path, format="markdown", category="markdown", metadata={"fixture": index})
        for index, path in enumerate(paths)
    ]
    main_vm = window._view_model
    main_vm.set_mode("batch")
    main_vm._files = refs
    window._sync_files_from_main_vm(refs)
    batch_vm = window._batch_list_vm
    assert batch_vm.get_files() == list(paths)
    assert len(window._file_contexts) == 2
    for index, source in enumerate(paths):
        assert window._batch_list.select_file(source)
        window._on_selected_file_changed(source)
        assert main_vm.selected_file is refs[index]
        builder = ExecutionRequestBuilder(
            main_vm, batch_vm, file_contexts=lambda: window._file_contexts, selected_template=lambda: None
        )
        assert builder.file_ref(source).metadata["fixture"] == index
        assert builder.file_ref(source).path == source
    batch_vm.set_file_status(first, "failed", error_message="first only")
    first_entry = batch_vm.get_file_entry(first)
    second_entry = batch_vm.get_file_entry(second)
    assert first_entry is not None and second_entry is not None
    assert first_entry.status == "failed"
    assert second_entry.status == "pending"
    main_vm.reserve_execution_inputs("operation", (first,))
    assert not main_vm.can_remove_file(first)
    assert main_vm.can_remove_file(second)
    main_vm.release_execution_inputs("operation")
    history = TaskHistory()
    record = history.remember({"request_id": "operation", "file_paths": list(paths)})
    history.record("operation", first, "failed")
    history.record("operation", second, "completed")
    assert record.failed_paths == [first]
    assert len(record.outcomes) == 2


def test_posix_keys_through_selection_request_and_status(monkeypatch, main_window):
    _use_path_flavor(monkeypatch, posixpath, "/")
    # Window-only fixture exercises the real widgets/models without a worker.
    # Native POSIX filesystem coverage is separate below.
    _assert_distinct_gui_inputs(main_window, (r"/vault/a\b.md", "/vault/a/b.md"))


@pytest.mark.skipif(os.name == "nt", reason="literal backslash filenames require a POSIX filesystem")
def test_two_existing_posix_files_keep_their_own_gui_state(tmp_path, main_window):
    first = tmp_path / r"a\b.md"
    second = tmp_path / "a" / "b.md"
    second.parent.mkdir()
    first.write_text("# First", encoding="utf-8")
    second.write_text("# Second", encoding="utf-8")
    _assert_distinct_gui_inputs(main_window, (str(first), str(second)))
    assert first.read_text(encoding="utf-8") == "# First"
    assert second.read_text(encoding="utf-8") == "# Second"


def test_windows_widget_lookup_uses_shared_alias_policy(monkeypatch, qapp):
    _use_path_flavor(monkeypatch, ntpath, "\\")
    vm = BatchListViewModel()
    vm.add_files(
        ["C:/Vault/Note.md"],
        file_resolver=lambda _path: {"detected_format": "markdown", "workflow_category": "markdown"},
    )
    widget = BatchList(view_model=vm)
    try:
        assert widget.select_file(r"c:\vault\NOTE.md")
        assert widget.get_current_file() == "C:/Vault/Note.md"
    finally:
        widget.deleteLater()
    main_vm = MainWindowViewModel(controller=None)
    main_vm.reserve_execution_inputs("operation", ("C:/Vault/Note.md",))
    assert not main_vm.can_remove_file(r"c:\vault\NOTE.md")
    main_vm.release_execution_inputs("operation")
