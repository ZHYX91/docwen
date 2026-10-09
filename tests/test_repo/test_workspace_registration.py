"""Explicit registration and independent clone/worktree engineering boundaries."""

from __future__ import annotations

import json
import subprocess
from pathlib import Path

import pytest
from tools import workspace_root

pytestmark = pytest.mark.unit


def test_initialize_and_resolve_clone_without_private_readme(tmp_path: Path) -> None:
    repo = tmp_path / "ordinary-clone"
    (repo / ".git").mkdir(parents=True)
    root = workspace_root.initialize_workspace(tmp_path / ".workspace", [repo])
    assert workspace_root.resolve_workspace_root(repo, explicit=root) == root
    assert workspace_root.registered_repositories(root) == (repo,)
    assert not (root / "README.md").exists()
    with pytest.raises(ValueError, match="already_exists"):
        workspace_root.initialize_workspace(root, [repo])


@pytest.mark.parametrize("entry", ["../outside", ".workspace", "repos/../../outside"])
def test_registry_cannot_expand_cleanup_boundary(tmp_path: Path, entry: str) -> None:
    root = workspace_root.initialize_workspace(tmp_path / ".workspace", [])
    (root / "workspace.json").write_text(json.dumps({"schema": "docwen.workspace.v1", "repositories": [entry]}))
    assert not workspace_root.is_governed_workspace_root(root)
    with pytest.raises(ValueError):
        workspace_root.registered_repositories(root)


def test_worktree_resolves_git_owned_primary_checkout_workspace(tmp_path: Path) -> None:
    repo = tmp_path / "repos" / "docwen"
    repo.mkdir(parents=True)

    def git(*args: str) -> None:
        subprocess.run(["git", "-C", str(repo), *args], check=True, capture_output=True)

    git("init")
    git("-c", "user.name=Test", "-c", "user.email=test@example.invalid", "commit", "--allow-empty", "-m", "initial")
    root = workspace_root.initialize_workspace(tmp_path / ".workspace", [repo])
    worktree = tmp_path / "worktrees" / "feature"
    git("worktree", "add", "--detach", str(worktree))
    try:
        assert workspace_root.resolve_workspace_root(worktree, environment={}) == root
    finally:
        git("worktree", "remove", str(worktree))
