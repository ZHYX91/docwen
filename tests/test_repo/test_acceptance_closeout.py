from __future__ import annotations

import hashlib
import json
from functools import partial
from pathlib import Path

import pytest
from tools import acceptance_closeout, workspace_cleanup, workspace_root

pytestmark = pytest.mark.unit


def _workspace(tmp_path: Path) -> Path:
    workspace = tmp_path / ".workspace"
    workspace.mkdir()
    (workspace / "README.md").write_text("# DocWen 本地工作区\n", encoding="utf-8")
    for name in workspace_root._GOVERNANCE_DIRECTORIES:
        (workspace / name).mkdir()
    return workspace


def _leased_run(workspace: Path, *, state: str, pid: int = 999_999_999) -> Path:
    run = (workspace / "temp" / "acceptance-run").resolve()
    run.mkdir()
    (run / workspace_cleanup.LEASE_NAME).write_text(
        json.dumps(
            {
                "schemaVersion": 1,
                "owner": "docwen.tests.acceptance",
                "kind": "acceptance-run",
                "pid": pid,
                "createdAt": "2026-09-02T00:00:00Z",
                "state": state,
                "root": str(run),
            }
        )
        + "\n",
        encoding="utf-8",
    )
    (run / "summary.log").write_bytes(b"passed\n")
    return run


def test_closeout_writes_compact_receipt_and_removes_raw_run(tmp_path: Path) -> None:
    workspace = _workspace(tmp_path)
    run = _leased_run(workspace, state="completed-success")
    subject = b"candidate"

    receipt = acceptance_closeout.close_run(
        workspace_root=workspace,
        run_root=run,
        candidate_id="candidate-1",
        gate="office-host",
        subject_name="DocWen.zip",
        subject_sha256=hashlib.sha256(subject).hexdigest(),
        subject_bytes=len(subject),
        limitations=("WPS not selected",),
    )

    assert not run.exists()
    payload = json.loads(receipt.read_text(encoding="utf-8"))
    assert payload["schema"] == acceptance_closeout.RECEIPT_SCHEMA
    assert payload["candidateId"] == "candidate-1"
    assert payload["gate"] == "office-host"
    assert payload["result"] == "passed"
    assert payload["rawRunRemoved"] is True
    assert payload["limitations"] == ["WPS not selected"]
    assert list((workspace / "diagnostics").iterdir()) == []


def test_closeout_rejects_nonterminal_or_live_run(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    workspace = _workspace(tmp_path)
    run = _leased_run(workspace, state="active", pid=1234)
    monkeypatch.setattr(workspace_cleanup, "_process_alive", lambda pid: pid == 1234)

    with pytest.raises(acceptance_closeout.AcceptanceCloseoutError, match="run_not_success_terminal"):
        acceptance_closeout.close_run(
            workspace_root=workspace,
            run_root=run,
            candidate_id="candidate-1",
            gate="office-host",
            subject_name="DocWen.zip",
            subject_sha256="0" * 64,
            subject_bytes=0,
        )

    assert run.is_dir()


@pytest.mark.parametrize("result", ["failed", "superseded"])
def test_closeout_preserves_failure_state(tmp_path: Path, result: str) -> None:
    workspace = _workspace(tmp_path)
    run = _leased_run(workspace, state="retained-failure")
    close = partial(
        acceptance_closeout.close_run,
        workspace_root=workspace,
        run_root=run,
        candidate_id="candidate-2",
        gate="source",
        subject_name="candidate.zip",
        subject_sha256="0" * 64,
        subject_bytes=0,
        result=result,
    )
    with pytest.raises(acceptance_closeout.AcceptanceCloseoutError, match="reason_required"):
        close()
    assert run.exists()
    (run / "summary.log").write_bytes(b"Assertion failed: target changed\n")
    receipt = close(
        reason="Original assertion failed; replaced by candidate-3/source.", evidence_files=("summary.log",)
    )
    payload = json.loads(receipt.read_text())
    assert payload["result"] == result
    assert payload["originalLease"]["state"] == "retained-failure"
    assert payload["rawRunRemoved"]
    assert payload["evidence"][0]["text"] == "Assertion failed: target changed\n"


def test_closeout_cleanup_failure_preserves_receipt_and_raw_run(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    workspace = _workspace(tmp_path)
    run = _leased_run(workspace, state="retained-failure")

    def fail(*args: object, **kwargs: object) -> None:
        raise PermissionError("denied")

    monkeypatch.setattr(workspace_cleanup, "apply_saved_plan", fail)
    with pytest.raises(PermissionError, match="denied"):
        acceptance_closeout.close_run(
            workspace_root=workspace,
            run_root=run,
            candidate_id="failure",
            gate="test",
            subject_name="candidate.zip",
            subject_sha256="0" * 64,
            subject_bytes=0,
            result="failed",
            reason="original failure",
            evidence_files=("summary.log",),
        )
    payload = json.loads((workspace / "acceptance/failure--test.json").read_text())
    assert payload["result"] == "failed"
    assert payload["cleanupError"] == "PermissionError: denied"
    assert payload["rawRunRemoved"] is False
    assert run.exists()
    assert payload["evidence"][0]["text"] == "passed\n"


@pytest.mark.parametrize("relative", ["../secret.txt", ".env", "missing.log"])
def test_closeout_rejects_unsafe_evidence_before_cleanup(tmp_path: Path, relative: str) -> None:
    workspace = _workspace(tmp_path)
    run = _leased_run(workspace, state="completed-success")
    (run / ".env").write_text("secret")
    with pytest.raises(acceptance_closeout.AcceptanceCloseoutError):
        acceptance_closeout.close_run(
            workspace_root=workspace,
            run_root=run,
            candidate_id="invalid",
            gate="test",
            subject_name="candidate.zip",
            subject_sha256="0" * 64,
            subject_bytes=0,
            evidence_files=(relative,),
        )
    assert run.exists()
    assert not list((workspace / "acceptance").iterdir())
