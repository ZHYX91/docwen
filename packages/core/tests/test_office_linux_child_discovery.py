"""Bounded procfs discovery fixtures; signal permission remains kernel-owned."""

from __future__ import annotations

import time
from pathlib import Path

import pytest

from docwen_core import office_process

pytestmark = pytest.mark.integration


def _status(root: Path, name: str, *, parent: int, namespace_ids: tuple[int, ...]) -> None:
    directory = root / name
    directory.mkdir(parents=True, exist_ok=True)
    (directory / "status").write_text(
        f"Name:\tfixture\nPPid:\t{parent}\nNSpid:\t" + "\t".join(map(str, namespace_ids)) + "\n",
        encoding="ascii",
    )


@pytest.fixture
def proc_view(tmp_path, monkeypatch):
    root = tmp_path / "proc"
    task = root / "self/task/1001"
    task.mkdir(parents=True)
    _status(root, "self", parent=1000, namespace_ids=(1001, 7))
    _status(root, "1003", parent=1001, namespace_ids=(1003, 9))
    _status(root, "1004", parent=1001, namespace_ids=(1004, 10, 1))
    _status(root, "1005", parent=1002, namespace_ids=(1005, 11))
    monkeypatch.setattr(office_process, "_PROC_ROOT", root)
    monkeypatch.setattr(office_process, "_TASK_ROOT", task.parent)
    monkeypatch.setattr(office_process.os, "getpid", lambda: 7)
    return root, task


@pytest.mark.parametrize("contents,expected", [("1003 1004\n", {9, 10}), ("", set())])
def test_readable_children_interface_precedes_status_scan(proc_view, monkeypatch, contents, expected):
    _root, task = proc_view
    (task / "children").write_text(contents, encoding="ascii")

    def forbidden_fallback(*_args):
        raise AssertionError("A readable children interface must remain the preferred source")

    monkeypatch.setattr(office_process, "_status_direct_children", forbidden_fallback)
    assert office_process._direct_children() == expected


def test_persistent_missing_children_uses_only_matching_parent_status(proc_view):
    _root, task = proc_view
    assert task.is_dir() and not (task / "children").exists()
    assert office_process._direct_children() == {9, 10}


def test_disappearing_secondary_thread_is_distinct_from_missing_main_interface(proc_view, monkeypatch):
    _root, task = proc_view
    (task / "children").write_text("1003\n", encoding="ascii")
    secondary = task.parent / "1002"
    secondary.mkdir()
    actual_reader = office_process._proc_text

    def read(path):
        if path == secondary / "children":
            secondary.rmdir()
            raise FileNotFoundError("The secondary thread exited")
        return actual_reader(path)

    def forbidden_fallback(*_args):
        raise AssertionError("An exited secondary thread does not invalidate the readable main interface")

    monkeypatch.setattr(office_process, "_proc_text", read)
    monkeypatch.setattr(office_process, "_status_direct_children", forbidden_fallback)
    assert office_process._direct_children() == {9}


def test_missing_task_directory_uses_parent_status_with_same_namespace_map(proc_view):
    _root, task = proc_view
    task.rmdir()
    task.parent.rmdir()
    assert office_process._direct_children() == {9, 10}


def test_candidate_namespace_mismatch_is_not_guessed(proc_view):
    root, _task = proc_view
    _status(root, "1003", parent=1001, namespace_ids=(1999, 9))
    with pytest.raises(OSError, match="Cannot map"):
        office_process._direct_children()


def test_scan_and_text_budgets_remain_finite(proc_view, monkeypatch):
    root, _task = proc_view
    with pytest.raises(OSError, match="finite scan budget"):
        office_process._status_direct_children(1001, 1, time.monotonic() - 1)
    monkeypatch.setattr(office_process, "_DISCOVERY_MAX_ENTRIES", 1)
    with pytest.raises(OSError, match="finite scan budget"):
        office_process._status_direct_children(1001, 1, time.monotonic() + 1)
    monkeypatch.setattr(office_process, "_PROC_TEXT_MAX_CHARS", 8)
    with pytest.raises(OSError, match="finite text budget"):
        office_process._namespace_pids(root / "self/status")
