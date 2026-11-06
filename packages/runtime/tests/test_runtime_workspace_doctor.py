"""Runtime environment probes use actual configured services and never disclose paths."""

from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest

from docwen_runtime.adapters import RuntimePortAdapter
from docwen_runtime.output.finalizer import AtomicPublishUnavailable
from docwen_runtime.workspace.manager import WorkspaceManager

pytestmark = pytest.mark.integration


def _port(root: Path, finalizer) -> RuntimePortAdapter:
    fake_manager = SimpleNamespace(_workspaces=WorkspaceManager(root), _finalizer=finalizer)
    return RuntimePortAdapter(task_manager=fake_manager)


def test_runtime_workspace_probe_writes_and_removes_only_private_probe(tmp_path: Path) -> None:
    root = tmp_path / "runtime"
    port = _port(root, SimpleNamespace())
    assert port.probe_runtime_workspace() == {"status": "ok", "reason": None}
    assert root.is_dir() and list(root.iterdir()) == []


def test_runtime_workspace_probe_denied_root_fails_without_replacing_it(tmp_path: Path) -> None:
    root = tmp_path / "occupied"
    root.write_text("original", encoding="utf-8")
    port = _port(root, SimpleNamespace())
    assert port.probe_runtime_workspace() == {"status": "fail", "reason": "not_writable"}
    assert root.read_text(encoding="utf-8") == "original"


def test_output_probe_uses_injected_finalizer_and_preserves_input(tmp_path: Path) -> None:
    chosen = tmp_path / "output"
    chosen.mkdir()
    source = chosen / "keep.docx"
    source.write_bytes(b"original")
    calls = []

    class Finalizer:
        def preflight(self, policy):
            calls.append(policy.output_dir)

    port = _port(tmp_path / "runtime", Finalizer())
    assert port.probe_output_directory(str(chosen)) == {"status": "ok", "reason": None}
    assert calls == [str(chosen)]
    assert source.read_bytes() == b"original"


def test_output_probe_rejects_relative_and_reports_atomic_unsupported(tmp_path: Path) -> None:
    class Finalizer:
        def preflight(self, policy):
            raise AtomicPublishUnavailable(95, "unsafe publication")

    port = _port(tmp_path / "runtime", Finalizer())
    assert port.probe_output_directory("relative/output")["reason"] == "invalid_path"
    assert port.probe_output_directory(str(tmp_path / "output")) == {
        "status": "fail",
        "reason": "atomic_publish_unsupported",
    }
