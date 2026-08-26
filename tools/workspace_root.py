from __future__ import annotations

import os
import stat
from collections.abc import Mapping
from pathlib import Path

WORKSPACE_ROOT_ENV = "DOCWEN_WORKSPACE_ROOT"
_GOVERNANCE_DIRECTORIES = (
    "acceptance",
    "artifacts",
    "backups",
    "cache",
    "diagnostics",
    "quarantine",
    "temp",
    "tools",
)
_README_HEADING = "# DocWen 本地工作区"


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
        raise WorkspaceRootError(f"unsupported_repository_layout:{repository}")
    engineering_root = repository.parent.parent
    candidate = engineering_root / ".workspace"
    if not is_governed_workspace_root(candidate):
        raise WorkspaceRootError(f"governed_workspace_root_not_found:{repository}")
    return candidate
