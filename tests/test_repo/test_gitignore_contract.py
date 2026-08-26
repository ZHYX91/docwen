"""Repository ignore rules must not hide unknown configuration files."""

from __future__ import annotations

import subprocess
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit


def test_gitignore_uses_explicit_dotfile_patterns() -> None:
    root = Path(__file__).resolve().parents[2]
    patterns = {
        line.strip()
        for line in (root / ".gitignore").read_text(encoding="utf-8").splitlines()
        if line.strip() and not line.lstrip().startswith("#")
    }

    assert ".*" not in patterns
    assert {".env", ".env.*", "!.env.example", "**/.venv/", "**/.pytest_cache*/"} <= patterns


def test_generated_acceptance_evidence_is_explicitly_ignored() -> None:
    root = Path(__file__).resolve().parents[2]
    patterns = {
        line.strip()
        for line in (root / ".gitignore").read_text(encoding="utf-8").splitlines()
        if line.strip() and not line.lstrip().startswith("#")
    }

    assert "/.acceptance-runtime/" in patterns


def test_tracked_files_do_not_embed_maintainer_workspace_roots() -> None:
    root = Path(__file__).resolve().parents[2]
    tracked = subprocess.run(
        ["git", "ls-files", "-z"],
        cwd=root,
        check=True,
        capture_output=True,
    ).stdout.split(b"\0")
    forbidden = (
        ("D:" + r"\Projects\DocWen-Workspace").encode(),
        ("C:" + r"\Users\zheng").encode(),
    )
    offenders: list[str] = []
    for raw_path in tracked:
        if not raw_path:
            continue
        path = root / raw_path.decode("utf-8", errors="strict")
        if path.is_file() and any(root_bytes in path.read_bytes() for root_bytes in forbidden):
            offenders.append(path.relative_to(root).as_posix())

    assert not offenders, f"tracked files expose maintainer workspace roots: {offenders}"
