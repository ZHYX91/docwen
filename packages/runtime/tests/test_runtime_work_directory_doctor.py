"""Runtime's doctor probe verifies the exact root without path leakage."""

from __future__ import annotations

from types import SimpleNamespace

import pytest

from docwen_runtime.adapters import RuntimePortAdapter

pytestmark = pytest.mark.unit


def test_runtime_probe_creates_then_removes_its_private_entry(tmp_path) -> None:
    root = tmp_path / "runtime"
    adapter = RuntimePortAdapter(task_manager=SimpleNamespace(workspace_root=root))
    assert adapter.probe_work_directory() is True
    assert root.is_dir()
    assert list(root.iterdir()) == []


def test_runtime_probe_reports_blocked_root_without_guessing_other_location(tmp_path) -> None:
    root = tmp_path / "blocked"
    root.write_text("this is a file", encoding="utf-8")
    adapter = RuntimePortAdapter(task_manager=SimpleNamespace(workspace_root=root))
    assert adapter.probe_work_directory() is False
    assert root.read_text(encoding="utf-8") == "this is a file"
