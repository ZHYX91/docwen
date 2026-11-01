"""Registered runtime roots retain repository and cleanup ownership boundaries."""

from __future__ import annotations

import json
import subprocess
from pathlib import Path

import pytest
from tools import workspace_cleanup, workspace_root

pytestmark = pytest.mark.unit


def _repository(root: Path, name: str = "docwen") -> Path:
    repo = root / "repos" / name
    (repo / ".git").mkdir(parents=True)
    return repo


def test_runtime_registry_and_discovery_are_independent_of_private_infrastructure(tmp_path: Path) -> None:
    repo = _repository(tmp_path)
    runtime = workspace_root.initialize_workspace(tmp_path / "runtime", [repo])
    assert json.loads((runtime / "workspace.json").read_text())["schema"] == "docwen.workspace.v2"
    assert workspace_root.resolve_workspace_root(repo, environment={}) == runtime
    assert workspace_root.registered_repositories(runtime) == (repo,)
    assert not (tmp_path / ".workspace").exists()
    assert not (tmp_path / "infra").exists()


@pytest.mark.parametrize("entry", ["../ObsidianPlugins/repos/plugin", "runtime", "runtime/temp", "repos/../../foreign"])
def test_runtime_registry_cannot_grant_cross_owner_or_runtime_source_ownership(tmp_path: Path, entry: str) -> None:
    runtime = workspace_root.initialize_workspace(tmp_path / "runtime", [])
    (runtime / "workspace.json").write_text(json.dumps({"schema": "docwen.workspace.v2", "repositories": [entry]}))
    with pytest.raises(workspace_root.WorkspaceRootError):
        workspace_root.registered_repositories(runtime)
    assert not workspace_root.is_governed_workspace_root(runtime)


def test_invalid_runtime_does_not_silently_fall_back_to_legacy_owner(tmp_path: Path) -> None:
    repo = _repository(tmp_path)
    workspace_root.initialize_workspace(tmp_path / ".workspace", [repo])
    (tmp_path / "runtime").mkdir()
    with pytest.raises(workspace_root.WorkspaceRootError, match="invalid_governed_workspace_root"):
        workspace_root.resolve_workspace_root(repo, environment={})


def test_registry_schema_is_bound_to_the_declared_directory_kind(tmp_path: Path) -> None:
    runtime = workspace_root.initialize_workspace(tmp_path / "runtime", [])
    (runtime / "workspace.json").write_text(json.dumps({"schema": "docwen.workspace.v1", "repositories": []}))
    assert not workspace_root.is_governed_workspace_root(runtime)


def test_nested_registered_runtime_does_not_become_parent_cleanup_authority(tmp_path: Path) -> None:
    runtime = workspace_root.initialize_workspace(tmp_path / "runtime", [])
    nested = runtime / "temp" / "fixture" / "runtime"
    nested.parent.mkdir(parents=True)
    workspace_root.initialize_workspace(nested, [])
    (nested / "temp" / "owned").mkdir()
    (nested / "temp" / "owned" / workspace_cleanup.LEASE_NAME).write_text("{}")
    errors: list[dict[str, str]] = []
    assert workspace_cleanup._lease_markers(runtime / "temp", errors=errors) == []
    assert any(item["reason"] == "nested_workspace_boundary" for item in errors)


def test_linked_worktree_uses_primary_registered_runtime(tmp_path: Path) -> None:
    repo = tmp_path / "repos" / "docwen"
    repo.mkdir(parents=True)

    def git(*args: str) -> None:
        subprocess.run(["git", "-C", str(repo), *args], check=True, capture_output=True)

    git("init")
    git("-c", "user.name=Test", "-c", "user.email=test@example.invalid", "commit", "--allow-empty", "-m", "initial")
    runtime = workspace_root.initialize_workspace(tmp_path / "runtime", [repo])
    worktree = tmp_path / "worktrees" / "feature"
    git("worktree", "add", "--detach", str(worktree))
    assert workspace_root.resolve_workspace_root(worktree, environment={}) == runtime
