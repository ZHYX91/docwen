from __future__ import annotations

from pathlib import Path

import pytest
from tools import workspace_root

pytestmark = pytest.mark.unit


def _governance_root(engineering_root: Path) -> Path:
    governed = engineering_root / ".workspace"
    governed.mkdir(parents=True)
    (governed / "README.md").write_text("# DocWen 本地工作区\n", encoding="utf-8")
    for name in workspace_root._GOVERNANCE_DIRECTORIES:
        (governed / name).mkdir()
    return governed


def test_resolver_supports_the_current_repos_layout(tmp_path: Path) -> None:
    engineering_root = tmp_path / "DocWen-Workspace"
    repo = engineering_root / "repos" / "docwen"
    repo.mkdir(parents=True)
    governed = _governance_root(engineering_root)

    assert workspace_root.resolve_workspace_root(repo, environment={}) == governed


def test_explicit_path_precedes_environment_and_discovery(tmp_path: Path) -> None:
    repo = tmp_path / "engineering" / "repos" / "docwen"
    repo.mkdir(parents=True)
    discovered = _governance_root(tmp_path / "engineering")
    configured_governed = _governance_root(tmp_path / "configured-root")
    explicit_governed = _governance_root(tmp_path / "explicit-root")

    assert (
        workspace_root.resolve_workspace_root(
            repo,
            environment={workspace_root.WORKSPACE_ROOT_ENV: str(configured_governed)},
        )
        == configured_governed
    )
    assert (
        workspace_root.resolve_workspace_root(
            repo,
            explicit=explicit_governed,
            environment={workspace_root.WORKSPACE_ROOT_ENV: str(configured_governed)},
        )
        == explicit_governed
    )
    assert discovered != configured_governed


def test_resolver_rejects_the_removed_direct_repository_layout(tmp_path: Path) -> None:
    engineering_root = tmp_path / "DocWen-Workspace"
    repo = engineering_root / "docwen"
    repo.mkdir(parents=True)
    _governance_root(engineering_root)

    with pytest.raises(workspace_root.WorkspaceRootError, match="unsupported_repository_layout"):
        workspace_root.resolve_workspace_root(repo, environment={})


@pytest.mark.parametrize("source", ["explicit", "environment"])
def test_resolver_rejects_ungoverned_configured_roots(tmp_path: Path, source: str) -> None:
    repo = tmp_path / "engineering" / "repos" / "docwen"
    repo.mkdir(parents=True)
    ungoverned = tmp_path / "ungoverned"
    ungoverned.mkdir()

    with pytest.raises(workspace_root.WorkspaceRootError, match="invalid_governed_workspace_root"):
        workspace_root.resolve_workspace_root(
            repo,
            explicit=ungoverned if source == "explicit" else None,
            environment={workspace_root.WORKSPACE_ROOT_ENV: str(ungoverned)} if source == "environment" else {},
        )


def test_discovery_fails_closed_without_creating_a_workspace(tmp_path: Path) -> None:
    repo = tmp_path / "engineering" / "repos" / "docwen"
    repo.mkdir(parents=True)

    with pytest.raises(workspace_root.WorkspaceRootError, match="governed_workspace_root_not_found"):
        workspace_root.resolve_workspace_root(repo, environment={})

    assert not (repo.parent / ".workspace").exists()
    assert not (repo.parents[1] / ".workspace").exists()


def test_repos_layout_ignores_an_unrelated_workspace_inside_the_repos_container(tmp_path: Path) -> None:
    engineering_root = tmp_path / "engineering"
    repo = engineering_root / "repos" / "docwen"
    repo.mkdir(parents=True)
    governed = _governance_root(engineering_root)
    _governance_root(engineering_root / "repos")

    assert workspace_root.resolve_workspace_root(repo, environment={}) == governed
