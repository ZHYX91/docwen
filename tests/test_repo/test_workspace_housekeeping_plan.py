from __future__ import annotations

import json
import os
import shutil
import subprocess
from datetime import UTC, datetime, timedelta
from pathlib import Path

import pytest
from tools import qa, workspace_cleanup, workspace_root

pytestmark = pytest.mark.unit


def _workspace(tmp_path: Path) -> tuple[Path, Path]:
    engineering = tmp_path / "DocWen-Workspace"
    workspace = engineering / ".workspace"
    workspace.mkdir(parents=True)
    (workspace / "README.md").write_text("# DocWen 本地工作区\n", encoding="utf-8")
    for name in workspace_root._GOVERNANCE_DIRECTORIES:
        (workspace / name).mkdir()
    (workspace / "build").mkdir()
    (workspace / "tmp").mkdir()
    (engineering / "repos" / "docwen").mkdir(parents=True)
    return engineering, workspace


def _lease(
    root: Path,
    *,
    created_at: datetime,
    state: str,
    kind: str = "pytest-runtime",
    pid: int = 999_999_999,
) -> None:
    root.mkdir(parents=True)
    (root / workspace_cleanup.LEASE_NAME).write_text(
        json.dumps(
            {
                "schemaVersion": 1,
                "owner": "docwen.tools.qa",
                "kind": kind,
                "pid": pid,
                "createdAt": created_at.isoformat().replace("+00:00", "Z"),
                "state": state,
                "root": str(root.resolve()),
            }
        )
        + "\n",
        encoding="utf-8",
    )


def test_retention_policy_cleans_success_immediately_and_bounds_failures() -> None:
    recent = timedelta(hours=1)

    assert workspace_cleanup.retention_decision(
        state="completed-success",
        age=timedelta(0),
        same_kind_rank=0,
    ) == {"eligible": True, "reason": "success_scratch"}
    assert workspace_cleanup.retention_decision(
        state="retained-failure",
        age=recent,
        same_kind_rank=0,
    ) == {"eligible": False, "reason": "failure_retained"}
    assert workspace_cleanup.retention_decision(
        state="retained-failure",
        age=recent,
        same_kind_rank=2,
    ) == {"eligible": True, "reason": "failure_retention_cap"}
    assert workspace_cleanup.retention_decision(
        state="retained-failure",
        age=timedelta(hours=72),
        same_kind_rank=0,
    ) == {"eligible": True, "reason": "failure_ttl_expired"}
    assert workspace_cleanup.retention_decision(
        state="retained-manual",
        age=timedelta(days=30),
        same_kind_rank=20,
    ) == {"eligible": False, "reason": "state_not_terminal"}


def test_default_plan_covers_managed_roots_and_only_reports_root_bypasses(tmp_path: Path) -> None:
    engineering, workspace = _workspace(tmp_path)
    now = datetime(2026, 8, 31, tzinfo=UTC)
    success = workspace / "build" / "success"
    old_failure = workspace / "tmp" / "old-failure"
    recent_failures = [workspace / "temp" / f"failure-{index}" for index in range(3)]
    _lease(success, created_at=now - timedelta(minutes=1), state="completed-success", kind="build")
    _lease(old_failure, created_at=now - timedelta(days=4), state="retained-failure", kind="old")
    for index, root in enumerate(recent_failures):
        _lease(
            root,
            created_at=now - timedelta(hours=index + 1),
            state="retained-failure",
            kind="same-kind",
        )
    bypass = engineering / "temp-manual-probe"
    bypass.mkdir()

    plan = workspace_cleanup.create_plan(workspace_root=workspace, now=now)

    planned = {Path(entry["path"]) for entry in plan["entries"]}
    assert success.resolve() in planned
    assert old_failure.resolve() in planned
    assert recent_failures[2].resolve() in planned
    assert recent_failures[0].resolve() not in planned
    assert recent_failures[1].resolve() not in planned
    assert bypass.resolve() not in planned
    assert plan["observations"]["rootTempBypasses"] == [
        {
            "path": str(bypass.resolve()),
            "isDirectory": True,
            "isReparse": False,
            "rootMtimeNs": bypass.stat().st_mtime_ns,
            "reportedOnly": True,
        }
    ]


def test_explicit_plan_can_include_managed_and_reported_bypass_targets(tmp_path: Path) -> None:
    engineering, workspace = _workspace(tmp_path)
    managed = workspace / "temp" / "old-probe"
    bypass = engineering / "temp-old-probe"
    managed.mkdir()
    bypass.mkdir()

    plan = workspace_cleanup.create_plan(
        workspace_root=workspace,
        explicit_targets=(managed, bypass),
        reason="superseded scratch",
    )

    assert [entry["source"] for entry in plan["entries"]] == ["explicit", "explicit-bypass"]
    assert {Path(entry["path"]) for entry in plan["entries"]} == {managed.resolve(), bypass.resolve()}


def test_apply_requires_a_saved_unchanged_plan_and_preflights_all_entries(tmp_path: Path) -> None:
    _, workspace = _workspace(tmp_path)
    first = workspace / "temp" / "first"
    second = workspace / "temp" / "second"
    first.mkdir()
    second.mkdir()
    (first / "payload.txt").write_text("first", encoding="utf-8")
    (second / "payload.txt").write_text("second", encoding="utf-8")
    plan = workspace_cleanup.create_plan(
        workspace_root=workspace,
        explicit_targets=(first, second),
        reason="test scratch",
    )
    plan_path = workspace_cleanup.save_plan(plan, workspace / "diagnostics" / "plan.json")
    (second / "payload.txt").write_text("changed", encoding="utf-8")

    with pytest.raises(workspace_cleanup.HousekeepingError, match="target_identity_changed"):
        workspace_cleanup.apply_saved_plan(plan_path, workspace_root=workspace)

    assert first.is_dir()
    assert second.is_dir()


def test_apply_rejects_same_content_directory_replacement(tmp_path: Path) -> None:
    _, workspace = _workspace(tmp_path)
    target = workspace / "temp" / "replaceable"
    target.mkdir()
    payload = target / "payload.txt"
    payload.write_text("same bytes", encoding="utf-8")
    target_mtime = target.stat().st_mtime_ns
    payload_mtime = payload.stat().st_mtime_ns
    plan = workspace_cleanup.create_plan(
        workspace_root=workspace,
        explicit_targets=(target,),
        reason="test scratch",
    )
    plan_path = workspace_cleanup.save_plan(plan, workspace / "diagnostics" / "plan.json")

    shutil.rmtree(target)
    target.mkdir()
    payload = target / "payload.txt"
    payload.write_text("same bytes", encoding="utf-8")
    os.utime(payload, ns=(payload_mtime, payload_mtime))
    os.utime(target, ns=(target_mtime, target_mtime))

    with pytest.raises(workspace_cleanup.HousekeepingError, match="target_identity_changed"):
        workspace_cleanup.apply_saved_plan(plan_path, workspace_root=workspace)

    assert target.is_dir()


def test_saved_plan_applies_after_per_target_revalidation(tmp_path: Path) -> None:
    _, workspace = _workspace(tmp_path)
    target = workspace / "temp" / "disposable"
    target.mkdir()
    (target / "payload.txt").write_text("payload", encoding="utf-8")
    plan = workspace_cleanup.create_plan(
        workspace_root=workspace,
        explicit_targets=(target,),
        reason="test scratch",
    )
    plan_path = workspace_cleanup.save_plan(plan, workspace / "diagnostics" / "plan.json")

    result = workspace_cleanup.apply_saved_plan(plan_path, workspace_root=workspace)

    assert result["removed"] == [str(target.resolve())]
    assert result["removedEntries"] == [{"path": str(target.resolve()), "bytes": len("payload")}]
    assert result["removedBytes"] == len("payload")
    assert not target.exists()
    assert plan_path.is_file()


def test_qa_uses_shared_retention_policy_and_removes_its_saved_plan(tmp_path: Path) -> None:
    _, workspace = _workspace(tmp_path)
    now = datetime.now(UTC)
    success = workspace / "temp" / "success"
    manual = workspace / "temp" / "manual"
    failures = [workspace / "temp" / f"failure-{index}" for index in range(3)]
    _lease(success, created_at=now, state="completed-success", kind="qa")
    _lease(manual, created_at=now - timedelta(days=30), state="retained-manual", kind="qa-manual")
    for index, root in enumerate(failures):
        _lease(
            root,
            created_at=now - timedelta(hours=index + 1),
            state="retained-failure",
            kind="qa-failure",
        )

    qa._cleanup_expired_workspace_temps(workspace)

    assert not success.exists()
    assert manual.is_dir()
    assert failures[0].is_dir()
    assert failures[1].is_dir()
    assert not failures[2].exists()
    assert list((workspace / "diagnostics").glob("qa-housekeeping-*.json")) == []


def test_plan_content_tampering_is_rejected(tmp_path: Path) -> None:
    _, workspace = _workspace(tmp_path)
    target = workspace / "temp" / "disposable"
    target.mkdir()
    plan = workspace_cleanup.create_plan(
        workspace_root=workspace,
        explicit_targets=(target,),
        reason="test scratch",
    )
    plan_path = workspace_cleanup.save_plan(plan, workspace / "diagnostics" / "plan.json")
    payload = json.loads(plan_path.read_text(encoding="utf-8"))
    payload["entries"][0]["reason"] = "tampered"
    plan_path.write_text(json.dumps(payload), encoding="utf-8")

    with pytest.raises(workspace_cleanup.HousekeepingError, match="plan_fingerprint_mismatch"):
        workspace_cleanup.apply_saved_plan(plan_path, workspace_root=workspace)

    assert target.is_dir()


def test_live_lease_blocks_explicit_plan(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    _, workspace = _workspace(tmp_path)
    target = workspace / "temp" / "live"
    _lease(
        target,
        created_at=datetime(2026, 8, 31, tzinfo=UTC),
        state="retained-failure",
        pid=4321,
    )
    monkeypatch.setattr(workspace_cleanup, "_process_alive", lambda pid: pid == 4321)

    with pytest.raises(workspace_cleanup.HousekeepingError, match="lease_process_alive"):
        workspace_cleanup.create_plan(
            workspace_root=workspace,
            explicit_targets=(target,),
            reason="must not delete live lease",
        )


def test_protected_and_repository_targets_are_not_normal_housekeeping(tmp_path: Path) -> None:
    engineering, workspace = _workspace(tmp_path)
    acceptance = workspace / "acceptance" / "receipt"
    acceptance.mkdir()
    repository = engineering / "repos" / "docwen"
    dependency = repository / ".venv"
    dependency.mkdir()

    with pytest.raises(workspace_cleanup.HousekeepingError, match="protected_target"):
        workspace_cleanup.create_plan(
            workspace_root=workspace,
            explicit_targets=(acceptance,),
            reason="forbidden",
        )
    with pytest.raises(workspace_cleanup.HousekeepingError, match="repository_target_requires_clean_deps"):
        workspace_cleanup.create_plan(
            workspace_root=workspace,
            explicit_targets=(repository,),
            reason="forbidden",
        )

    plan = workspace_cleanup.create_plan(
        workspace_root=workspace,
        clean_dependency_targets=(dependency,),
    )
    assert plan["entries"][0]["source"] == "clean-deps"
    assert plan["entries"][0]["path"] == str(dependency.resolve())

    with pytest.raises(workspace_cleanup.HousekeepingError, match="clean_deps_name_required"):
        workspace_cleanup.create_plan(
            workspace_root=workspace,
            clean_dependency_targets=(repository / "src",),
        )


def test_managed_dependency_requires_explicit_clean_deps_and_is_not_auto_planned(tmp_path: Path) -> None:
    _, workspace = _workspace(tmp_path)
    dependency = workspace / "tmp" / "probe" / "node_modules"
    _lease(
        dependency,
        created_at=datetime(2020, 1, 1, tzinfo=UTC),
        state="completed-success",
        kind="dependency-probe",
    )

    default_plan = workspace_cleanup.create_plan(
        workspace_root=workspace,
        now=datetime(2026, 8, 31, tzinfo=UTC),
    )

    assert all(Path(entry["path"]) != dependency.resolve() for entry in default_plan["entries"])
    assert {
        "path": str(dependency.resolve()),
        "reason": "dependency_target_requires_clean_deps",
    } in default_plan["observations"]["skipped"]
    with pytest.raises(workspace_cleanup.HousekeepingError, match="dependency_target_requires_clean_deps"):
        workspace_cleanup.create_plan(
            workspace_root=workspace,
            explicit_targets=(dependency,),
            reason="not sufficient",
        )

    explicit_plan = workspace_cleanup.create_plan(
        workspace_root=workspace,
        clean_dependency_targets=(dependency,),
    )
    assert explicit_plan["entries"][0]["source"] == "clean-deps"


def test_reparse_target_outside_planned_root_is_rejected(tmp_path: Path) -> None:
    _, workspace = _workspace(tmp_path)
    target = workspace / "temp" / "linked"
    outside = workspace / "temp" / "outside"
    target.mkdir()
    outside.mkdir()
    link = target / "escape"
    try:
        os.symlink(outside, link, target_is_directory=True)
    except (NotImplementedError, OSError) as error:
        if os.name != "nt":
            pytest.skip(f"directory symlink unavailable: {error}")
        junction = subprocess.run(
            ["cmd.exe", "/d", "/c", "mklink", "/J", str(link), str(outside)],
            check=False,
            capture_output=True,
            text=True,
        )
        if junction.returncode != 0:
            pytest.skip(f"directory junction unavailable: {junction.stderr or junction.stdout}")

    with pytest.raises(workspace_cleanup.HousekeepingError, match="reparse_target_outside_target"):
        workspace_cleanup.create_plan(
            workspace_root=workspace,
            explicit_targets=(target,),
            reason="unsafe link",
        )


def test_plan_must_be_saved_under_workspace_diagnostics(tmp_path: Path) -> None:
    _, workspace = _workspace(tmp_path)
    target = workspace / "temp" / "disposable"
    target.mkdir()
    plan = workspace_cleanup.create_plan(
        workspace_root=workspace,
        explicit_targets=(target,),
        reason="test scratch",
    )

    with pytest.raises(workspace_cleanup.HousekeepingError, match="plan_path_must_be_under_diagnostics"):
        workspace_cleanup.save_plan(plan, tmp_path / "plan.json")


def test_cli_separates_saved_plan_generation_from_apply(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    _, workspace = _workspace(tmp_path)
    target = workspace / "temp" / "cli-disposable"
    target.mkdir()
    (target / "payload.txt").write_text("payload", encoding="utf-8")
    plan_path = workspace / "diagnostics" / "cli-plan.json"

    assert (
        workspace_cleanup.main(
            [
                "--workspace-root",
                str(workspace),
                "--target",
                str(target),
                "--reason",
                "cli test scratch",
                "--plan-output",
                str(plan_path),
            ]
        )
        == 0
    )
    preview = json.loads(capsys.readouterr().out)
    assert preview["savedPlanPath"] == str(plan_path.resolve())
    assert target.is_dir()

    assert (
        workspace_cleanup.main(
            [
                "--workspace-root",
                str(workspace),
                "--apply-plan",
                str(plan_path),
            ]
        )
        == 0
    )
    applied = json.loads(capsys.readouterr().out)
    assert applied["removed"] == [str(target.resolve())]
    assert applied["removedBytes"] == len("payload")
    assert not target.exists()
