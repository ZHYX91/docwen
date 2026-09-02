from __future__ import annotations

import hashlib
import json
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
    (run / "summary.log").write_text("passed\n", encoding="utf-8")
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
