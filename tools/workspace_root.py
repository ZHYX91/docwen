from __future__ import annotations

import argparse
import json
import os
import stat
import subprocess
from collections.abc import Mapping
from pathlib import Path

WORKSPACE_ROOT_ENV = "DOCWEN_WORKSPACE_ROOT"
_GOVERNANCE_DIRECTORIES = (
    "acceptance",
    "artifacts",
    "backups",
    "build",
    "cache",
    "diagnostics",
    "quarantine",
    "temp",
    "tools",
)
_README_HEADING = "# DocWen 本地工作区"
REGISTRY_NAME = "workspace.json"


class WorkspaceRootError(ValueError):
    """A governed DocWen workspace root could not be selected safely."""


def _absolute(path: Path) -> Path:
    return Path(os.path.abspath(os.fspath(path)))


def _is_reparse(metadata: os.stat_result) -> bool:
    flag = int(getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0))
    return bool(flag and int(getattr(metadata, "st_file_attributes", 0)) & flag)


def _is_plain_directory(path: Path) -> bool:
    try:
        metadata = path.lstat()
    except OSError:
        return False
    return stat.S_ISDIR(metadata.st_mode) and not stat.S_ISLNK(metadata.st_mode) and not _is_reparse(metadata)


def is_governed_workspace_root(path: Path) -> bool:
    """Return whether *path* carries the existing DocWen governance contract."""
    root = _absolute(path)
    if not _is_plain_directory(root):
        return False
    if (root / REGISTRY_NAME).exists() or (root / REGISTRY_NAME).is_symlink():
        try:
            registered_repositories(root)
        except (OSError, ValueError):
            return False
        return all(_is_plain_directory(root / name) for name in _GOVERNANCE_DIRECTORIES)
    readme = root / "README.md"
    try:
        metadata = readme.lstat()
        heading_matches = readme.read_text(encoding="utf-8").startswith(_README_HEADING)
    except OSError:
        return False
    if not stat.S_ISREG(metadata.st_mode) or stat.S_ISLNK(metadata.st_mode) or _is_reparse(metadata):
        return False
    return heading_matches and all(_is_plain_directory(root / name) for name in _GOVERNANCE_DIRECTORIES)


def resolve_workspace_root(
    repo_root: Path,
    *,
    explicit: Path | None = None,
    environment: Mapping[str, str] | None = None,
) -> Path:
    """Resolve the governed root without creating an inferred directory."""
    if explicit is not None:
        selected = _absolute(explicit)
        if not is_governed_workspace_root(selected):
            raise WorkspaceRootError(f"invalid_governed_workspace_root:{selected}")
        return selected

    values = os.environ if environment is None else environment
    configured = values.get(WORKSPACE_ROOT_ENV, "").strip()
    if configured:
        selected = _absolute(Path(configured))
        if not is_governed_workspace_root(selected):
            raise WorkspaceRootError(f"invalid_governed_workspace_root:{selected}")
        return selected

    repository = _absolute(repo_root)
    if repository.parent.name.casefold() != "repos":
        # A worktree shares its registered checkout's Git directory. This is a
        # single Git-owned identity lookup, not an arbitrary ancestor search.
        try:
            completed = subprocess.run(
                ["git", "-C", str(repository), "rev-parse", "--path-format=absolute", "--git-common-dir"],
                capture_output=True,
                text=True,
                check=False,
                timeout=10,
            )
        except (OSError, subprocess.TimeoutExpired) as exc:
            # Hermetic builds deliberately exclude Git from PATH. Discovery is
            # unavailable there; callers must keep their existing owned runtime.
            raise WorkspaceRootError(f"repository_identity_unavailable:{repository}") from exc
        common = Path(completed.stdout.strip())
        if completed.returncode or common.name != ".git" or common.parent.parent.name.casefold() != "repos":
            raise WorkspaceRootError(f"unsupported_repository_layout:{repository}")
        repository = common.parent
    engineering_root = repository.parent.parent
    candidate = engineering_root / ".workspace"
    if not is_governed_workspace_root(candidate):
        raise WorkspaceRootError(f"governed_workspace_root_not_found:{repository}")
    return candidate


def registered_repositories(workspace: Path) -> tuple[Path, ...]:
    """Read explicit local registration; an absent registry grants no repo cleanup."""
    marker = workspace / REGISTRY_NAME
    if not marker.exists() and not marker.is_symlink():
        return ()
    metadata = marker.lstat()
    if not stat.S_ISREG(metadata.st_mode) or _is_reparse(metadata):
        raise WorkspaceRootError("unsafe_workspace_registry")
    payload = json.loads(marker.read_text(encoding="utf-8"))
    if not isinstance(payload, dict) or payload.get("schema") != "docwen.workspace.v1":
        raise WorkspaceRootError("invalid_workspace_registry")
    entries = payload.get("repositories")
    if not isinstance(entries, list) or any(not isinstance(item, str) for item in entries):
        raise WorkspaceRootError("invalid_registered_repositories")
    result = []
    for entry in entries:
        relative = Path(entry)
        if relative.is_absolute() or not relative.parts or any(part in {".", ".."} for part in relative.parts):
            raise WorkspaceRootError("invalid_registered_repository_path")
        path = workspace.parent / relative
        if path == workspace or workspace in path.parents:
            raise WorkspaceRootError("workspace_cannot_be_registered_repository")
        for current in (path, *path.parents):
            if current == workspace.parent:
                break
            if (current.exists() or current.is_symlink()) and not _is_plain_directory(current):
                raise WorkspaceRootError("unsafe_registered_repository")
        if path in result:
            raise WorkspaceRootError("duplicate_registered_repository")
        result.append(path)
    return tuple(result)


def initialize_workspace(root: Path, repositories: list[Path]) -> Path:
    """Initialize only an explicitly selected new root or existing legacy root."""
    root = _absolute(root)
    if root.name != ".workspace" or not root.parent.is_dir():
        raise WorkspaceRootError("explicit_workspace_directory_required")
    if any(not _is_plain_directory(parent) for parent in (root.parent, *root.parent.parents)):
        raise WorkspaceRootError("unsafe_workspace_parent")
    if root.exists() and not is_governed_workspace_root(root):
        raise WorkspaceRootError("refuse_existing_ungoverned_workspace")
    if (root / REGISTRY_NAME).exists():
        raise WorkspaceRootError("workspace_registry_already_exists")
    relatives = []
    for repository in repositories:
        repository = _absolute(repository)
        if (
            not repository.is_relative_to(root.parent)
            or repository == root.parent
            or root.is_relative_to(repository)
            or repository.is_relative_to(root)
        ):
            raise WorkspaceRootError("repository_must_be_inside_engineering_root")
        if not _is_plain_directory(repository) or not (repository / ".git").exists():
            raise WorkspaceRootError("registered_repository_must_be_git_checkout")
        if any(
            not _is_plain_directory(parent)
            for parent in repository.parents
            if parent != root.parent and parent.is_relative_to(root.parent)
        ):
            raise WorkspaceRootError("unsafe_registered_repository")
        relative = repository.relative_to(root.parent).as_posix()
        if relative in relatives:
            raise WorkspaceRootError("duplicate_registered_repository")
        relatives.append(relative)
    root.mkdir(exist_ok=True)
    for name in _GOVERNANCE_DIRECTORIES:
        (root / name).mkdir(exist_ok=True)
    with (root / REGISTRY_NAME).open("x", encoding="utf-8", newline="\n") as stream:
        json.dump({"schema": "docwen.workspace.v1", "repositories": relatives}, stream, indent=2)
        stream.write("\n")
    return root


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Initialize an explicit engineering workspace; never scan or delete.")
    parser.add_argument("--init", type=Path, required=True)
    parser.add_argument("--repository", type=Path, action="append", default=[])
    arguments = parser.parse_args()
    print(initialize_workspace(arguments.init, arguments.repository))
