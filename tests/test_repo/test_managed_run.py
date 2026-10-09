"""Engineering entry failures remain discoverable without losing their cause."""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest
from tools import run_lease, run_source, workspace_root

pytestmark = pytest.mark.unit


def _managed(parent: Path):
    return run_lease.managed_run(parent, prefix="probe-", owner="docwen.test", kind="fault-injection")


@pytest.mark.parametrize("phase", ["identity", "marker"])
def test_initialization_failure_removes_only_its_empty_owned_run(tmp_path, monkeypatch, phase):
    def fail(*args, **kwargs):
        raise OSError("injected initialization failure")

    monkeypatch.setattr(
        run_lease if phase == "identity" else run_lease.json, "lease_payload" if phase == "identity" else "dump", fail
    )
    with pytest.raises(OSError, match="injected initialization"), _managed(tmp_path):
        pytest.fail("Initialization should not yield")
    assert list(tmp_path.iterdir()) == []


def test_initialization_does_not_remove_unexpected_entries(tmp_path, monkeypatch, capsys):
    def fail(root, **kwargs):
        (root / "foreign").write_bytes(b"keep")
        raise OSError("identity unavailable")

    monkeypatch.setattr(run_lease, "lease_payload", fail)
    with pytest.raises(OSError, match="identity unavailable"), _managed(tmp_path):
        pass
    root = next(tmp_path.iterdir())
    assert (root / "foreign").read_bytes() == b"keep"
    assert str(root) in capsys.readouterr().err


@pytest.mark.parametrize(
    "failure,state", [(RuntimeError("body"), "retained-failure"), (KeyboardInterrupt(), "retained-interrupted")]
)
def test_body_failure_records_terminal_outcome(tmp_path, failure, state):
    with pytest.raises(type(failure)) as caught, _managed(tmp_path) as run:
        raise failure
    assert caught.value is failure
    lease = json.loads((run.root / ".docwen-temp-lease.json").read_text())
    assert lease["state"] == lease["outcome"] == state


def test_failed_terminal_write_preserves_original_exception_and_valid_active_lease(tmp_path, monkeypatch):
    primary = RuntimeError("conversion failed")
    with pytest.raises(RuntimeError) as caught, _managed(tmp_path) as run:

        def fail(*args, **kwargs):
            raise OSError("disk full")

        monkeypatch.setattr(run_lease.json, "dump", fail)
        raise primary
    assert caught.value is primary
    assert "disk full" in str(primary.__notes__)
    assert json.loads((run.root / ".docwen-temp-lease.json").read_text())["state"] == "active"


@pytest.mark.parametrize("remove_marker", [False, True])
def test_success_cleanup_failure_keeps_success_outcome(tmp_path, monkeypatch, remove_marker):
    def fail(root, *args, **kwargs):
        if remove_marker:
            (root / ".docwen-temp-lease.json").unlink()
        raise OSError("busy")

    monkeypatch.setattr(run_lease.shutil, "rmtree", fail)
    with pytest.raises(OSError, match="busy"), _managed(tmp_path) as run:
        pass
    lease = json.loads((run.root / ".docwen-temp-lease.json").read_text())
    assert lease["state"] == "retained-cleanup-failure"
    assert lease["outcome"] == "completed-success"


def test_cleanup_does_not_overwrite_replaced_marker_and_leaves_external_recovery(tmp_path, monkeypatch):
    def fail(root, *args, **kwargs):
        marker = root / ".docwen-temp-lease.json"
        marker.rename(root / "old-marker")
        marker.write_text('{"owner":"other"}')
        raise OSError("cleanup failed")

    monkeypatch.setattr(run_lease.shutil, "rmtree", fail)
    with pytest.raises(OSError, match="cleanup failed"), _managed(tmp_path) as run:
        pass
    assert json.loads((run.root / ".docwen-temp-lease.json").read_text()) == {"owner": "other"}
    record = json.loads((tmp_path / f"{run.root.name}.recovery.json").read_text())
    assert record["rootIdentity"]["inode"] == run.root.stat().st_ino
    assert record["lease"]["outcome"] == "completed-success"
    assert record["phase"] == "cleanup-recording"


def test_success_terminal_write_failure_keeps_active_lease_and_external_outcome(tmp_path, monkeypatch):
    original_replace = Path.replace

    def fail(self, target):
        if self.name == ".docwen-temp-lease.next":
            raise OSError("terminal write failed")
        return original_replace(self, target)

    monkeypatch.setattr(Path, "replace", fail)
    with pytest.raises(OSError, match="terminal write failed"), _managed(tmp_path) as run:
        pass
    assert json.loads((run.root / ".docwen-temp-lease.json").read_text())["state"] == "active"
    record = json.loads((tmp_path / f"{run.root.name}.recovery.json").read_text())
    assert record["lease"]["outcome"] == "completed-success"
    assert record["phase"] == "terminal-recording"


@pytest.mark.parametrize("returncode", [0, 7])
def test_source_launcher_preserves_child_exit_status_and_lease(tmp_path, monkeypatch, returncode):
    (tmp_path / "temp").mkdir()
    monkeypatch.setattr(run_source, "resolve_workspace_root", lambda *args, **kwargs: tmp_path)

    def execute(*args, **kwargs):
        assert Path(kwargs["env"]["TEMP"]).is_dir()
        return SimpleNamespace(returncode=returncode)

    monkeypatch.setattr(run_source.subprocess, "run", execute)
    assert run_source.main(["cli", "--", "--version"]) == returncode
    roots = list((tmp_path / "temp").iterdir())
    if returncode:
        assert json.loads((roots[0] / ".docwen-temp-lease.json").read_text())["state"] == "retained-failure"
    else:
        assert roots == []


@pytest.mark.parametrize("phase", ["config", "qt", "evidence"])
def test_render_initialization_and_evidence_failures_remain_leased(tmp_path, monkeypatch, phase):
    import builtins

    from scripts.maintenance import render_screenshots

    repo = tmp_path / "repo"
    repo.mkdir()
    (repo / "configs").mkdir()
    workspace = tmp_path / "workspace"
    (workspace / "temp").mkdir(parents=True)
    monkeypatch.setattr(workspace_root, "resolve_workspace_root", lambda *args, **kwargs: workspace)
    monkeypatch.setattr(os, "environ", dict(os.environ))
    monkeypatch.setattr(sys, "path", sys.path.copy())
    monkeypatch.chdir(repo)
    if phase != "config":
        (repo / "configs/gui.toml").write_text('locale = "en_US"\n')
    if phase == "qt":
        original_import = builtins.__import__

        def import_without_qt(name, *args, **kwargs):
            if name.startswith("PySide6"):
                raise ImportError("Qt unavailable")
            return original_import(name, *args, **kwargs)

        monkeypatch.setattr(builtins, "__import__", import_without_qt)
    if phase == "evidence":

        def fail_gui(*args):
            raise RuntimeError("original render failure")

        monkeypatch.setitem(sys.modules, "docwen_bundle", SimpleNamespace(gui_entry=SimpleNamespace(main=fail_gui)))
        monkeypatch.setitem(
            sys.modules, "docwen_gui", SimpleNamespace(app=SimpleNamespace(create_main_window=lambda: None))
        )
        monkeypatch.setattr(
            render_screenshots.subprocess,
            "check_output",
            lambda args, **kwargs: "commit" if kwargs.get("text") else b"",
        )
        original_write = Path.write_text

        def write_without_evidence(self, text, *args, **kwargs):
            if self.name == "en_US-light.json":
                raise OSError("evidence disk full")
            return original_write(self, text, *args, **kwargs)

        monkeypatch.setattr(Path, "write_text", write_without_evidence)
    error = {"config": FileNotFoundError, "qt": ImportError, "evidence": RuntimeError}[phase]
    with pytest.raises(error) as caught:
        render_screenshots.main(["--repo", str(repo), "en_US", "light", "--output", str(tmp_path / "media")])
    root = next((workspace / "temp").iterdir())
    assert json.loads((root / ".docwen-temp-lease.json").read_text())["state"] == "retained-failure"
    if phase == "evidence":
        assert "original render failure" in str(caught.value)
        assert "evidence disk full" in str(caught.value.__notes__)
