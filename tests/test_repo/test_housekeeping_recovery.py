from __future__ import annotations

import json
import os
from datetime import UTC, datetime, timedelta
from pathlib import Path

import pytest
from tools import housekeeping_journal, run_lease, workspace_cleanup

pytestmark = pytest.mark.unit


def _plan(tmp_path: Path) -> tuple[Path, Path, list[Path]]:
    workspace = tmp_path / ".workspace"
    (workspace / "diagnostics").mkdir(parents=True)
    roots = [workspace / "temp" / name for name in ("first", "second")]
    for root in roots:
        root.mkdir(parents=True)
        (root / "input.txt").write_text("controlled scratch")
    plan = workspace_cleanup.create_plan(
        workspace_root=workspace, explicit_targets=roots, reason="controlled test", disposition="delete"
    )
    return workspace, workspace_cleanup.save_plan(plan, workspace / "diagnostics" / "plan.json"), roots


def test_completed_plan_is_repeatable_without_deleting_recreated_paths(tmp_path: Path) -> None:
    workspace, saved, roots = _plan(tmp_path)
    first = workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace)
    assert workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace) == first
    roots[0].mkdir()
    (roots[0] / "new.txt").write_text("must survive")
    with pytest.raises(ValueError, match="completed_target_changed"):
        workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace)
    assert (roots[0] / "new.txt").read_text() == "must survive"


def test_restart_between_targets_uses_durable_completed_receipt(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    workspace, saved, roots = _plan(tmp_path)
    original = workspace_cleanup._revalidate_entry
    count = 0

    def interrupt(entry, *, workspace):
        nonlocal count
        count += 1
        if count == 4:  # Both preflight checks, first apply, then second apply.
            raise KeyboardInterrupt
        return original(entry, workspace=workspace)

    monkeypatch.setattr(workspace_cleanup, "_revalidate_entry", interrupt)
    with pytest.raises(KeyboardInterrupt):
        workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace)
    assert not roots[0].exists() and roots[1].exists()
    progress = json.loads(saved.with_suffix(".progress.json").read_text())
    assert [item["path"] for item in progress["completed"]] == [str(roots[0])]
    monkeypatch.setattr(workspace_cleanup, "_revalidate_entry", original)
    result = workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace)
    assert result["removed"] == [str(root) for root in roots]


def test_refused_mutation_is_recorded_and_never_retried(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    workspace, saved, roots = _plan(tmp_path)
    calls = []

    def refuse(*args, **kwargs):
        calls.append(args)
        raise PermissionError("controlled refusal")

    monkeypatch.setattr(workspace_cleanup.shutil, "rmtree", refuse)
    with pytest.raises(PermissionError, match="controlled refusal"):
        workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace)
    with pytest.raises(ValueError, match="previous_apply_outcome_unconfirmed"):
        workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace)
    assert len(calls) == 1 and all(root.exists() for root in roots)
    progress = json.loads(saved.with_suffix(".progress.json").read_text())
    assert "PermissionError" in progress["pending"]["error"]


def test_concurrent_apply_is_refused_without_removing_any_target(tmp_path: Path) -> None:
    workspace, saved, roots = _plan(tmp_path)
    with (
        housekeeping_journal.exclusive_apply(workspace / "diagnostics" / "housekeeping-apply.lock"),
        pytest.raises(OSError),
    ):
        workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace)
    assert all(root.exists() for root in roots)


def test_manual_retention_deadline_requests_review_but_never_selects_for_removal(tmp_path: Path) -> None:
    workspace = tmp_path / ".workspace"
    root = workspace / "temp" / "held"
    root.mkdir(parents=True)
    payload = run_lease.lease_payload(root, owner="docwen.test", kind="scratch")
    run_lease.transition(payload, root=root, owner="docwen.test", state="completed-success")
    run_lease.transition(payload, root=root, owner="docwen.test", state="retained-manual")
    payload["pid"] = 999999999
    payload.pop("processIdentity", None)
    (root / workspace_cleanup.LEASE_NAME).write_text(json.dumps(payload))
    plan = workspace_cleanup.create_plan(workspace_root=workspace, now=datetime.now(UTC) + timedelta(days=4))
    assert not plan["entries"]
    assert plan["observations"]["skipped"][0]["reason"] == "manual_retention_review_due"
    assert "does not authorize" in plan["observations"]["skipped"][0]["action"]
    assert payload["outcome"] == "completed-success"


def test_cleanup_failure_preserves_successful_run_outcome(tmp_path: Path) -> None:
    payload = run_lease.lease_payload(tmp_path, owner="docwen.test", kind="scratch")
    run_lease.transition(payload, root=tmp_path, owner="docwen.test", state="completed-success")
    run_lease.transition(payload, root=tmp_path, owner="docwen.test", state="retained-cleanup-failure")
    assert payload["outcome"] == "completed-success"
    assert payload["state"] == "retained-cleanup-failure"
    assert payload["finishedAt"]


def test_fresh_automatic_scan_does_not_bypass_recorded_refusal(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    workspace, saved, roots = _plan(tmp_path)
    # Add a dead successful lease before taking the new saved identity.
    payload = run_lease.lease_payload(roots[0], owner="docwen.test", kind="scratch")
    payload.update(pid=999999999, state="completed-success")
    payload.pop("processIdentity", None)
    (roots[0] / workspace_cleanup.LEASE_NAME).write_text(json.dumps(payload))
    plan = workspace_cleanup.create_plan(
        workspace_root=workspace, explicit_targets=[roots[0]], reason="controlled refusal", disposition="delete"
    )
    saved = workspace_cleanup.save_plan(plan, saved.with_name("refused.json"))

    def refuse(*args, **kwargs):
        raise PermissionError("controlled refusal")

    monkeypatch.setattr(workspace_cleanup.shutil, "rmtree", refuse)
    with pytest.raises(PermissionError):
        workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace)
    preview = workspace_cleanup.create_plan(workspace_root=workspace)
    assert not preview["entries"]
    assert preview["observations"]["skipped"][0]["reason"] == "previous_apply_outcome_unconfirmed"


def test_failed_terminal_recording_holds_expired_active_run(tmp_path, monkeypatch):
    workspace = tmp_path / ".workspace"
    parent = workspace / "temp"
    parent.mkdir(parents=True)
    original_replace = Path.replace

    def fail(self, target):
        if self.name == ".docwen-temp-lease.next":
            raise OSError("terminal write failed")
        return original_replace(self, target)

    with monkeypatch.context() as patch:
        patch.setattr(Path, "replace", fail)
        with (
            pytest.raises(OSError, match="terminal write failed"),
            run_lease.managed_run(parent, prefix="held-", owner="docwen.test", kind="scratch") as run,
        ):
            pass
    marker = run.root / workspace_cleanup.LEASE_NAME
    assert json.loads(marker.read_text())["state"] == "active"
    monkeypatch.setattr(workspace_cleanup, "_lease_process_alive", lambda payload: False)
    plan = workspace_cleanup.create_plan(workspace_root=workspace, now=datetime.now(UTC) + timedelta(days=4))
    assert not plan["entries"]
    assert "pending_run_recovery" in plan["observations"]["skipped"][0]["reason"]
    assert marker.exists()


@pytest.mark.parametrize("record_form", ["success", "malformed", "directory", "unreadable"])
def test_recovery_hold_rechecked_when_applying_saved_plan(tmp_path, monkeypatch, record_form):
    workspace, saved, roots = _plan(tmp_path)
    recovery = roots[0].parent / f"{roots[0].name}.recovery.json"
    if record_form == "directory":
        recovery.mkdir()
    else:
        recovery.write_text('{"knownOutcome":"completed-success"}' if record_form == "success" else "invalid")
    if record_form == "unreadable":
        original_lstat = Path.lstat

        def refuse(self, *args, **kwargs):
            if self == recovery:
                raise PermissionError("record inaccessible")
            return original_lstat(self, *args, **kwargs)

        monkeypatch.setattr(Path, "lstat", refuse)
    with pytest.raises(workspace_cleanup.HousekeepingError, match="run_recovery"):
        workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace)
    with pytest.raises(workspace_cleanup.HousekeepingError, match="run_recovery"):
        workspace_cleanup.create_plan(workspace_root=workspace, explicit_targets=[roots[0]], reason="manual target")
    assert all((root / "input.txt").read_text() == "controlled scratch" for root in roots)


@pytest.mark.parametrize("marker_present", [True, False])
@pytest.mark.parametrize("record_form", ["malformed", "directory"])
def test_parent_cleanup_holds_child_recovery(tmp_path, marker_present, record_form):
    workspace = tmp_path / ".workspace"
    (workspace / "diagnostics").mkdir(parents=True)
    group = workspace / "temp" / "group"
    child = group / "child-run"
    child.mkdir(parents=True)
    (child / "output.txt").write_text("preserve result")
    if marker_present:
        payload = run_lease.lease_payload(child, owner="docwen.test", kind="scratch")
        payload.update(pid=999999999, state="completed-success")
        payload.pop("processIdentity", None)
        (child / workspace_cleanup.LEASE_NAME).write_text(json.dumps(payload))
    plan = workspace_cleanup.create_plan(
        workspace_root=workspace, explicit_targets=[group], reason="parent cleanup", disposition="delete"
    )
    saved = workspace_cleanup.save_plan(plan, workspace / "diagnostics" / "parent.json")
    recovery = group / "child-run.recovery.json"
    if record_form == "directory":
        recovery.mkdir()
    else:
        recovery.write_text("invalid")
    with pytest.raises(workspace_cleanup.HousekeepingError, match="pending_run_recovery"):
        workspace_cleanup.create_plan(workspace_root=workspace, explicit_targets=[group], reason="parent cleanup")
    with pytest.raises(workspace_cleanup.HousekeepingError, match="pending_run_recovery"):
        workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace)
    assert (child / "output.txt").read_text() == "preserve result"
    assert (child / workspace_cleanup.LEASE_NAME).exists() == marker_present
    assert recovery.exists()


@pytest.mark.parametrize("marker_form", ["symlink", "directory", "foreign"])
def test_invalid_parent_marker_cannot_hide_child_recovery(tmp_path, marker_form):
    workspace, saved, roots = _plan(tmp_path)
    marker = roots[0] / workspace_cleanup.LEASE_NAME
    if marker_form == "symlink":
        try:
            marker.symlink_to(roots[0] / "input.txt")
        except OSError as error:
            pytest.skip(f"symlink creation unavailable: {error}")
    elif marker_form == "directory":
        marker.mkdir()
    else:
        marker.write_text(json.dumps({"schemaVersion": 1, "owner": "foreign"}))
    recovery = roots[0] / "child-run.recovery.json"
    recovery.write_text("unreconciled")
    with pytest.raises(workspace_cleanup.HousekeepingError, match=r"unsafe_lease_marker|foreign_lease_owner"):
        workspace_cleanup.create_plan(workspace_root=workspace, explicit_targets=[roots[0]], reason="parent cleanup")
    with pytest.raises(workspace_cleanup.HousekeepingError, match=r"unsafe_lease_marker|foreign_lease_owner"):
        workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace)
    assert recovery.read_text() == "unreconciled"
    assert all((root / "input.txt").read_text() == "controlled scratch" for root in roots)


def test_failed_atomic_lease_update_preserves_original_marker(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from scripts.release import build_production_candidate as production

    production._write_work_lease(tmp_path, state="active")
    marker = tmp_path / production.PRODUCTION_WORK_LEASE
    original = marker.read_bytes()

    def refuse(*args, **kwargs):
        raise OSError("controlled write failure")

    monkeypatch.setattr(production, "atomic_write", refuse)
    with pytest.raises(OSError):
        production._update_work_lease(tmp_path, state="retained-interrupted")
    assert marker.read_bytes() == original


def test_progress_write_failure_prevents_mutation(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    workspace, saved, roots = _plan(tmp_path)

    def refuse(self):
        raise OSError("cannot save progress")

    monkeypatch.setattr(housekeeping_journal.ApplyJournal, "save", refuse)
    with pytest.raises(OSError, match="cannot save progress"):
        workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace)
    assert all(root.exists() for root in roots)


@pytest.mark.skipif(os.name != "nt", reason="Windows recovery metadata")
def test_replay_requires_original_recycle_recovery(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from tools import recycle

    workspace, saved, roots = _plan(tmp_path)
    plan = workspace_cleanup.create_plan(
        workspace_root=workspace, explicit_targets=[roots[0]], reason="recovery verification"
    )
    assert plan["disposition"] == "recycle"
    saved = workspace_cleanup.save_plan(plan, saved.with_name("recycle.json"))
    result = workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace)
    assert result["removedEntries"][0]["recovery"]
    assert workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace) == result
    monkeypatch.setattr(recycle, "recovery_entries", lambda path: [])
    with pytest.raises(ValueError, match="completed_recovery_unverified"):
        workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace)


def test_journal_cannot_be_reused_with_another_plan(tmp_path: Path) -> None:
    workspace, saved, roots = _plan(tmp_path)
    journal = housekeeping_journal.ApplyJournal(saved.with_suffix(".progress.json"), "wrong-fingerprint")
    journal.save()
    with pytest.raises(ValueError, match="journal_identity_mismatch"):
        workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace)
    assert all(root.exists() for root in roots)
