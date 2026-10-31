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


def test_terminal_recording_rechecks_foreign_marker_after_preparing_replacement(tmp_path, monkeypatch):
    original_dump = run_lease.json.dump
    with pytest.raises(ValueError, match="marker_changed"), _managed(tmp_path) as run:

        def replace_marker_during_dump(payload, stream, *args, **kwargs):
            result = original_dump(payload, stream, *args, **kwargs)
            if Path(stream.name).name == ".docwen-temp-lease.next":
                marker = run.root / ".docwen-temp-lease.json"
                marker.rename(run.root / "original-marker")
                marker.write_bytes(b'{"owner":"foreign-during-save"}')
            return result

        monkeypatch.setattr(run_lease.json, "dump", replace_marker_during_dump)
    assert (run.root / ".docwen-temp-lease.json").read_bytes() == b'{"owner":"foreign-during-save"}'
    recovery = json.loads((tmp_path / f"{run.root.name}.recovery.json").read_text())
    assert recovery["phase"] == "terminal-recording"
    assert recovery["knownOutcome"] == "completed-success"


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


@pytest.mark.parametrize("body_result", ["success", "failure-state", "exception"])
def test_replaced_terminal_marker_preserves_known_outcome_and_both_errors(tmp_path, body_result):
    primary = RuntimeError("execution failed")
    expected_error = RuntimeError if body_result == "exception" else ValueError
    with pytest.raises(expected_error) as caught, _managed(tmp_path) as run:
        marker = run.root / ".docwen-temp-lease.json"
        marker.rename(run.root / "original-marker")
        marker.write_text('{"owner":"other"}')
        if body_result == "exception":
            raise primary
        if body_result == "failure-state":
            run.state = "retained-failure"
    record = json.loads((tmp_path / f"{run.root.name}.recovery.json").read_text())
    assert record["knownOutcome"] == ("completed-success" if body_result == "success" else "retained-failure")
    assert record["lease"]["state"] == "active"
    assert record["recordingError"]["type"] == "ValueError"
    assert "managed_run_marker_changed" in record["recordingError"]["message"]
    assert json.loads(marker.read_text()) == {"owner": "other"}
    if body_result == "exception":
        assert caught.value is primary
        assert record["executionError"] == {"type": "RuntimeError", "message": "execution failed"}
    else:
        assert record["executionError"] is None


@pytest.mark.parametrize("returncode", [0, 7])
@pytest.mark.parametrize("entry", ["cli", "gui"])
def test_source_launcher_preserves_child_exit_status_and_lease(tmp_path, monkeypatch, returncode, entry):
    (tmp_path / "temp").mkdir()
    monkeypatch.setattr(run_source, "resolve_workspace_root", lambda *args, **kwargs: tmp_path)

    def execute(*args, **kwargs):
        assert Path(kwargs["env"]["TEMP"]).is_dir()
        assert args[0][2] == ("docwen_gui" if entry == "gui" else "docwen_bundle.cli_entry")
        return SimpleNamespace(returncode=returncode)

    monkeypatch.setattr(run_source.subprocess, "run", execute)
    assert run_source.main([entry, "--", "--version"]) == returncode
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


def test_root_pin_allows_real_cleanup_and_closes_before_success_returns(tmp_path):
    with _managed(tmp_path) as run:
        (run.root / "owned-sample").write_bytes(b"finished sample")
    # On Windows RemoveDirectory can be pending until the directory pin closes.
    assert not run.root.exists()
    assert list(tmp_path.iterdir()) == []


@pytest.mark.skipif(os.name != "posix", reason="POSIX unlink/recreate identity boundary; native Windows is separate")
def test_cleanup_root_pin_rejects_deleted_and_recreated_root_without_adopting_it(tmp_path, monkeypatch):
    original_identity = []

    def replace_root_then_fail(root, *args, **kwargs):
        original_identity.append((root.stat().st_dev, root.stat().st_ino))
        for child in root.iterdir():
            child.unlink()
        root.rmdir()
        root.mkdir()
        (root / "foreign-owner").write_bytes(b"must remain untouched")
        assert (root.stat().st_dev, root.stat().st_ino) != original_identity[0]
        raise OSError("partial cleanup removed and recreated the root")

    monkeypatch.setattr(run_lease.shutil, "rmtree", replace_root_then_fail)
    with pytest.raises(OSError, match="partial cleanup removed and recreated"), _managed(tmp_path) as run:
        (run.root / "owned-sample").write_bytes(b"original sample")
    assert (run.root / "foreign-owner").read_bytes() == b"must remain untouched"
    assert not (run.root / ".docwen-temp-lease.json").exists()
    record = json.loads((tmp_path / f"{run.root.name}.recovery.json").read_text())
    assert record["rootIdentity"] == {"device": original_identity[0][0], "inode": original_identity[0][1]}
    assert record["phase"] == "cleanup-recording"
    assert record["knownOutcome"] == "completed-success"
    assert "managed_run_identity_changed" in record["recordingError"]["message"]


def test_source_environment_preserves_parent_values_and_prioritizes_all_checkout_packages(tmp_path, monkeypatch):
    repo = tmp_path / "current"
    cli = repo / "packages/apps/cli/src"
    optimizer = repo / "packages/plugins/optimizers/gongwen/src"
    for path in (cli, optimizer):
        path.mkdir(parents=True)
    old = str(tmp_path / "other-checkout/src")
    monkeypatch.setenv("PYTHONPATH", old)
    monkeypatch.setenv("UNRELATED_CALLER_VALUE", "preserve")
    before = dict(os.environ)
    run = tmp_path / "owned-run"
    env = run_source.source_environment(repo, run)
    assert env["PYTHONPATH"].split(os.pathsep) == [str(cli), str(optimizer), old]
    assert env["UNRELATED_CALLER_VALUE"] == "preserve"
    assert env["DOCWEN_RUNTIME_ROOT"] == str(run / "runtime")
    assert env["TEMP"] == env["TMP"] == env["TMPDIR"] == str(run / "temp")
    assert dict(os.environ) == before
