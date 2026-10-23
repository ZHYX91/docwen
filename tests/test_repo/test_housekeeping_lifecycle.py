from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path

import pytest
from scripts.release.publication_contract import CHECKSUM_NAME, package_names
from tools import qa, recycle, run_lease, workspace_cleanup

pytestmark = pytest.mark.unit


def test_unknown_leases_are_reported_even_without_cleanup_entries(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    workspace = tmp_path / ".workspace"
    root = workspace / "temp" / "legacy"
    root.mkdir(parents=True)
    (root / workspace_cleanup.LEASE_NAME).write_text(json.dumps({"schemaVersion": 1, "owner": "codex"}))
    qa._cleanup_expired_workspace_temps(workspace)
    assert "foreign_owner" in capsys.readouterr().err
    assert root.exists()


def test_run_lease_rejects_foreign_identity_and_unknown_state(tmp_path: Path) -> None:
    with pytest.raises(ValueError):
        run_lease.lease_payload(tmp_path, owner="codex", kind="native")
    payload = run_lease.lease_payload(tmp_path, owner="docwen.acceptance", kind="native")
    with pytest.raises(ValueError):
        run_lease.transition(payload, root=tmp_path, owner="docwen.other", state="retained-failure")
    with pytest.raises(ValueError):
        run_lease.transition(payload, root=tmp_path, owner="docwen.acceptance", state="blocked")
    run_lease.transition(payload, root=tmp_path, owner="docwen.acceptance", state="retained-interrupted")
    assert payload["state"] == "retained-interrupted"


@pytest.mark.skipif(os.name != "nt", reason="Windows recycle mode")
def test_recycle_failure_does_not_fall_back_to_delete(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    workspace = tmp_path / ".workspace"
    root = workspace / "temp" / "old"
    root.mkdir(parents=True)
    diagnostics = workspace / "diagnostics"
    diagnostics.mkdir()
    (root / "keep.txt").write_text("original")
    plan = workspace_cleanup.create_plan(
        workspace_root=workspace, explicit_targets=[root], reason="test", disposition="recycle"
    )
    saved = workspace_cleanup.save_plan(plan, diagnostics / "plan.json")

    def fail(path: Path, **kwargs) -> None:
        raise OSError("recycle unavailable")

    monkeypatch.setattr(recycle, "recycle_directory", fail)
    with pytest.raises(OSError, match="recycle unavailable"):
        workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace)
    assert (root / "keep.txt").read_text() == "original"
    (root / "keep.txt").write_text("changed")
    with pytest.raises(workspace_cleanup.HousekeepingError, match="previous_apply_outcome_unconfirmed"):
        workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace)


def test_publication_retirement_requires_verified_receipt_and_exact_bytes(tmp_path: Path) -> None:
    workspace = tmp_path / ".workspace"
    candidate = workspace / "artifacts" / "candidate"
    candidate.mkdir(parents=True)
    (workspace / "diagnostics").mkdir()
    (workspace / "acceptance").mkdir()
    receipt = workspace / "acceptance" / "published.json"
    content = b"payload"
    digest = hashlib.sha256(content).hexdigest()
    version = "0.20.0"
    names = [*package_names(version), CHECKSUM_NAME]
    for name in [*names, "candidate.json"]:
        (candidate / name).write_bytes(content)
    state = {
        "schema": "docwen-publication-progress-v1",
        "version": version,
        "stage": "draft-verified",
        "pending": None,
        "provenance": "verified",
        "releaseId": 1,
        "manifestSha256": digest,
        "assets": {name: {"sha256": digest, "bytes": len(content)} for name in names},
    }
    receipt.write_text(json.dumps(state))
    with pytest.raises(workspace_cleanup.HousekeepingError, match="publication_not_verified"):
        workspace_cleanup.plan_published_candidate(candidate, receipt)
    state["stage"] = "verified"
    state["assets"]["DocWen-windows-x64.zip"]["bytes"] += 1
    receipt.write_text(json.dumps(state))
    with pytest.raises(workspace_cleanup.HousekeepingError, match="publication_asset_changed"):
        workspace_cleanup.plan_published_candidate(candidate, receipt)
    state["assets"]["DocWen-windows-x64.zip"]["bytes"] -= 1
    receipt.write_text(json.dumps(state))
    saved = workspace_cleanup.plan_published_candidate(candidate, receipt)
    assert saved is not None and candidate.exists()
    # A copied or edited verification receipt cannot authorize the saved plan.
    receipt.write_text(json.dumps(state) + "\n")
    with pytest.raises(workspace_cleanup.HousekeepingError, match="publication_receipt_changed"):
        workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace)
    (candidate / "DocWen-windows-x64.zip").write_bytes(b"other")
    with pytest.raises(workspace_cleanup.HousekeepingError, match="publication_asset_changed"):
        workspace_cleanup.plan_published_candidate(candidate, receipt)


@pytest.mark.parametrize("damage", ["missing-mac", "extra", "missing-version"])
def test_publication_retirement_rejects_incomplete_or_foreign_assets(tmp_path: Path, damage: str) -> None:
    workspace = tmp_path / ".workspace"
    candidate = workspace / "artifacts" / "candidate"
    candidate.mkdir(parents=True)
    (workspace / "acceptance").mkdir()
    receipt = workspace / "acceptance" / "published.json"
    names = [*package_names("0.20.0"), CHECKSUM_NAME]
    state = {
        "schema": "docwen-publication-progress-v1",
        "version": "0.20.0",
        "stage": "verified",
        "pending": None,
        "provenance": "verified",
        "releaseId": 1,
        "assets": {name: {} for name in names},
    }
    if damage == "missing-mac":
        del state["assets"]["DocWenCLI-0.20.0-macos-arm64.tar.gz"]
    elif damage == "extra":
        state["assets"]["unrelated.zip"] = {}
    else:
        del state["version"]
    receipt.write_text(json.dumps(state))
    with pytest.raises(workspace_cleanup.HousekeepingError, match="publication_inventory_incomplete"):
        workspace_cleanup.plan_published_candidate(candidate, receipt)


@pytest.mark.parametrize("change", ["missing", "interpreter", "prefix", "toolSha256"])
def test_apply_requires_the_recorded_execution_environment_before_any_mutation(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, change: str
) -> None:
    workspace = tmp_path / ".workspace"
    root = workspace / "temp" / "scratch"
    root.mkdir(parents=True)
    (root / "keep.txt").write_bytes(b"preserve")
    (workspace / "diagnostics").mkdir()
    plan = workspace_cleanup.create_plan(
        workspace_root=workspace, explicit_targets=[root], reason="environment fixture", disposition="delete"
    )
    if change == "missing":
        del plan["executionEnvironment"]
    else:
        plan["executionEnvironment"][change] = "different environment"
    plan["planFingerprint"] = workspace_cleanup._plan_fingerprint(plan)
    saved = workspace_cleanup.save_plan(plan, workspace / "diagnostics" / "environment.json")

    def unexpected(*args, **kwargs):
        pytest.fail("environment mismatch reached filesystem mutation")

    monkeypatch.setattr(workspace_cleanup.shutil, "rmtree", unexpected)
    with pytest.raises(ValueError, match="plan_execution_environment"):
        workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace)
    assert (root / "keep.txt").read_bytes() == b"preserve"
    assert not saved.with_suffix(".progress.json").exists()


def test_cleanup_result_keeps_the_same_execution_environment_as_the_plan(tmp_path: Path) -> None:
    workspace = tmp_path / ".workspace"
    root = workspace / "temp" / "scratch"
    root.mkdir(parents=True)
    (workspace / "diagnostics").mkdir()
    plan = workspace_cleanup.create_plan(
        workspace_root=workspace, explicit_targets=[root], reason="environment fixture", disposition="delete"
    )
    saved = workspace_cleanup.save_plan(plan, workspace / "diagnostics" / "environment.json")
    result = workspace_cleanup.apply_saved_plan(saved, workspace_root=workspace)
    assert result["executionEnvironment"] == plan["executionEnvironment"]
    assert Path(result["executionEnvironment"]["interpreter"]).is_absolute()
