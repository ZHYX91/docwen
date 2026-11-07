"""Runtime's doctor probe verifies the exact root without path leakage."""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen_runtime.adapters import RuntimePortAdapter
from docwen_runtime.engine.route_resolver import RouteResolver
from docwen_runtime.engine.task_manager import TaskManager
from docwen_runtime.output.finalizer import OutputFinalizer
from docwen_runtime.plugin_registry.registry import PluginRegistry
from docwen_runtime.workspace.manager import WorkspaceManager

pytestmark = pytest.mark.unit


def _adapter(workspace_root: Path) -> RuntimePortAdapter:
    registry = PluginRegistry()
    manager = TaskManager(
        plugin_registry=registry,
        route_resolver=RouteResolver(registry),
        workspace_manager=WorkspaceManager(workspace_root),
        output_finalizer=OutputFinalizer(),
    )
    return RuntimePortAdapter(task_manager=manager)


def test_runtime_probe_creates_then_removes_its_private_entry(tmp_path) -> None:
    root = tmp_path / "runtime"
    adapter = _adapter(root)
    assert adapter.probe_work_directory() is True
    assert root.is_dir()
    assert list(root.iterdir()) == []


def test_runtime_probe_reports_blocked_root_without_guessing_other_location(tmp_path) -> None:
    root = tmp_path / "blocked"
    root.write_text("this is a file", encoding="utf-8")
    adapter = _adapter(root)
    assert adapter.probe_work_directory() is False
    assert root.read_text(encoding="utf-8") == "this is a file"


def test_runtime_probe_reports_cleanup_failure(tmp_path, monkeypatch) -> None:
    root = tmp_path / "runtime"
    adapter = _adapter(root)

    def denied(_path):
        raise PermissionError(13, "probe cleanup denied")

    with monkeypatch.context() as patcher:
        patcher.setattr("docwen_core.tempfiles.shutil.rmtree", denied)
        assert adapter.probe_work_directory() is False
    assert len(list(root.glob(".__docwen-doctor-*"))) == 1


@pytest.mark.parametrize("relative", [".", "./output"])
def test_output_probe_accepts_relative_directories_like_conversion(tmp_path, monkeypatch, relative: str) -> None:
    monkeypatch.chdir(tmp_path)
    adapter = _adapter(tmp_path / "runtime")
    assert adapter.probe_output_directory(relative) == "ok"
    assert adapter.probe_output_directory(str(Path(relative).resolve())) == "ok"
    assert list(tmp_path.iterdir()) == []


def test_output_probe_reports_malformed_path_as_invalid(tmp_path) -> None:
    adapter = _adapter(tmp_path / "runtime")
    assert adapter.probe_output_directory("invalid\0path") == "invalid_path"


def test_output_probe_does_not_turn_internal_failures_into_permission_errors(tmp_path, monkeypatch) -> None:
    adapter = _adapter(tmp_path / "runtime")

    def broken_preflight(_self, _policy):
        raise RuntimeError("unexpected finalizer failure")

    monkeypatch.setattr(OutputFinalizer, "preflight", broken_preflight)
    with pytest.raises(RuntimeError, match="unexpected finalizer failure"):
        adapter.probe_output_directory(str(tmp_path))
