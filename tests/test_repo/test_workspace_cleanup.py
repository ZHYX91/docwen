from __future__ import annotations

import json
from datetime import UTC, datetime, timedelta
from pathlib import Path

import pytest
from tools import workspace_cleanup, workspace_root

pytestmark = pytest.mark.unit


def _governance_root(engineering_root: Path) -> Path:
    governed = engineering_root / ".workspace"
    governed.mkdir(parents=True)
    (governed / "README.md").write_text("# DocWen 本地工作区\n", encoding="utf-8")
    for name in workspace_root._GOVERNANCE_DIRECTORIES:
        (governed / name).mkdir()
    return governed


def _lease(root: Path, *, created_at: datetime, pid: int = 999_999_999) -> None:
    root.mkdir(parents=True)
    (root / workspace_cleanup.LEASE_NAME).write_text(
        json.dumps(
            {
                "schemaVersion": 1,
                "owner": "docwen.tools.qa",
                "kind": "pytest-runtime",
                "pid": pid,
                "createdAt": created_at.isoformat().replace("+00:00", "Z"),
                "state": "retained-failure",
                "root": str(root.resolve()),
            }
        )
        + "\n",
        encoding="utf-8",
    )


def test_cleanup_defaults_to_a_non_mutating_preview(tmp_path: Path) -> None:
    workspace = tmp_path / ".workspace"
    runtime = workspace / "temp" / "docwen" / "pytest" / "old"
    now = datetime(2026, 8, 22, tzinfo=UTC)
    _lease(runtime, created_at=now - timedelta(days=4))

    result = workspace_cleanup.cleanup(
        workspace_root=workspace,
        max_age=timedelta(hours=72),
        apply=False,
        now=now,
    )

    assert result["eligible"] == [str(runtime.resolve())]
    assert result["removed"] == []
    assert runtime.is_dir()


def test_cleanup_discovers_the_governance_root_above_repos(tmp_path: Path) -> None:
    engineering_root = tmp_path / "DocWen-Workspace"
    repo = engineering_root / "repos" / "docwen"
    repo.mkdir(parents=True)
    governed = _governance_root(engineering_root)

    assert workspace_root.resolve_workspace_root(repo, environment={}) == governed


def test_cleanup_direct_apply_is_disabled_and_preview_remains_non_mutating(tmp_path: Path) -> None:
    workspace = tmp_path / ".workspace"
    now = datetime(2026, 8, 22, tzinfo=UTC)
    expired = workspace / "temp" / "docwen" / "pytest" / "expired"
    recent = workspace / "temp" / "docwen" / "pytest" / "recent"
    unmarked = workspace / "temp" / "docwen" / "pytest" / "unmarked"
    _lease(expired, created_at=now - timedelta(days=4))
    _lease(recent, created_at=now - timedelta(hours=1))
    unmarked.mkdir(parents=True)

    with pytest.raises(workspace_cleanup.HousekeepingError, match="direct_apply_disabled_use_saved_plan"):
        workspace_cleanup.cleanup(
            workspace_root=workspace,
            max_age=timedelta(hours=72),
            apply=True,
            now=now,
        )

    result = workspace_cleanup.cleanup(
        workspace_root=workspace,
        max_age=timedelta(hours=72),
        apply=False,
        now=now,
    )

    assert result["eligible"] == [str(expired.resolve())]
    assert result["removed"] == []
    assert expired.is_dir()
    assert recent.is_dir()
    assert unmarked.is_dir()


def test_cleanup_rejects_a_foreign_marker(tmp_path: Path) -> None:
    workspace = tmp_path / ".workspace"
    runtime = workspace / "temp" / "docwen" / "pytest" / "foreign"
    now = datetime(2026, 8, 22, tzinfo=UTC)
    _lease(runtime, created_at=now - timedelta(days=4))
    marker = runtime / workspace_cleanup.LEASE_NAME
    payload = json.loads(marker.read_text(encoding="utf-8"))
    payload["owner"] = "some-other-tool"
    marker.write_text(json.dumps(payload), encoding="utf-8")

    result = workspace_cleanup.cleanup(
        workspace_root=workspace,
        max_age=timedelta(0),
        apply=False,
        now=now,
    )

    assert runtime.is_dir()
    assert result["removed"] == []
    assert result["skipped"][0]["reason"] == "foreign_owner"


def test_cleanup_treats_a_leased_runtime_as_one_root(tmp_path: Path) -> None:
    workspace = tmp_path / ".workspace"
    runtime = workspace / "temp" / "docwen" / "pytest" / "parent"
    nested = runtime / "basetemp" / "synthetic" / "nested"
    now = datetime(2026, 8, 22, tzinfo=UTC)
    _lease(runtime, created_at=now - timedelta(days=4))
    _lease(nested, created_at=now - timedelta(days=4))

    markers = workspace_cleanup._lease_markers(workspace / "temp" / "docwen")

    assert markers == [runtime / workspace_cleanup.LEASE_NAME]


def test_cleanup_does_not_treat_nested_workspace_fixtures_as_owned_roots(tmp_path: Path) -> None:
    workspace = tmp_path / ".workspace"
    fixture_runtime = (
        workspace
        / "temp"
        / "caller-owned-pytest-runtime"
        / "basetemp"
        / "test-case"
        / ".workspace"
        / "temp"
        / "docwen"
        / "pytest"
        / "expired-fixture"
    )
    now = datetime(2026, 8, 22, tzinfo=UTC)
    _lease(fixture_runtime, created_at=now - timedelta(days=4))

    result = workspace_cleanup.cleanup(
        workspace_root=workspace,
        max_age=timedelta(hours=72),
        apply=False,
        now=now,
    )

    assert result["eligible"] == []
    assert result["skipped"] == [{"path": str(fixture_runtime.parents[3]), "reason": "nested_workspace_boundary"}]
    assert fixture_runtime.is_dir()


def test_cleanup_unmounts_an_owned_short_drive_only_when_applying(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workspace = tmp_path / ".workspace"
    runtime = workspace / "temp" / "docwen" / "pytest" / "expired"
    now = datetime(2026, 8, 22, tzinfo=UTC)
    _lease(runtime, created_at=now - timedelta(days=4))
    marker = runtime / workspace_cleanup.LEASE_NAME
    payload = json.loads(marker.read_text(encoding="utf-8"))
    payload["shortDrive"] = "W:"
    marker.write_text(json.dumps(payload), encoding="utf-8")
    unmounted: list[tuple[str, Path]] = []
    monkeypatch.setattr(
        workspace_cleanup,
        "unmount_short_drive",
        lambda drive, *, expected_target: unmounted.append((drive, expected_target)),
    )

    preview = workspace_cleanup.cleanup(
        workspace_root=workspace,
        max_age=timedelta(0),
        apply=False,
        now=now,
    )

    assert preview["eligible"] == [str(runtime.resolve())]
    assert unmounted == []

    diagnostics = workspace / "diagnostics"
    diagnostics.mkdir()
    plan = workspace_cleanup.create_plan(
        workspace_root=workspace,
        explicit_targets=(runtime,),
        reason="expired test runtime",
        now=now,
    )
    plan_path = workspace_cleanup.save_plan(plan, diagnostics / "cleanup-plan.json")
    applied = workspace_cleanup.apply_saved_plan(plan_path, workspace_root=workspace)

    assert applied["removed"] == [str(runtime.resolve())]
    assert unmounted == [("W:", runtime.resolve())]
