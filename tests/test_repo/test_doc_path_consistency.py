"""Gate test: doc-path consistency checker (gate item 6 of 架构收口方案).

Asserts that ``tools/validation/check_doc_path_consistency.py``:
  1. Exists and exits 0 on the real repo (no unlabeled old-path references
     in current-architecture docs).
  2. Flags an old-path reference that lacks historical/migration context.
  3. Does NOT flag the same reference when the doc carries a historical
     context label, and does NOT flag whitelisted migration-mapping docs.

The rule is defined in ``docs/testing.md`` §4.1.
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

REPO_ROOT = Path(__file__).resolve().parents[2]
CHECKER = REPO_ROOT / "tools" / "validation" / "check_doc_path_consistency.py"


def _run_checker(repo_root: Path) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(CHECKER), "--repo-root", str(repo_root)],
        capture_output=True,
        text=True,
        timeout=60,
    )


def test_checker_exists() -> None:
    """The doc-path consistency checker script must exist (gate item 6)."""
    assert CHECKER.is_file(), f"Missing required gate script: {CHECKER}"


def test_checker_passes_on_real_repo() -> None:
    """The real repo must have no unlabeled old-path references in structural docs."""
    assert CHECKER.is_file(), f"Missing required gate script: {CHECKER}"
    result = _run_checker(REPO_ROOT)
    msg = f"exit={result.returncode}\nstdout:\n{result.stdout}\nstderr:\n{result.stderr}"
    assert result.returncode == 0, msg


def test_checker_flags_unlabeled_old_path(tmp_path: Path) -> None:
    """A current-architecture doc that references old paths without historical label must be flagged."""
    docs = tmp_path / "docs"
    arch = docs / "architecture"
    arch.mkdir(parents=True)
    (arch / "structure.md").write_text(
        "# GUI 结构\n\n组件位于 src/docwen/gui/core/window.py，使用 combobox_adapter 适配。\n",
        encoding="utf-8",
    )

    result = _run_checker(tmp_path)

    assert result.returncode != 0, "checker must flag unlabeled old-path references"
    assert "structure.md" in result.stdout


def test_checker_allows_labeled_historical_doc(tmp_path: Path) -> None:
    """A doc with a historical-context label may keep old paths (no false positive)."""
    docs = tmp_path / "docs"
    arch = docs / "architecture"
    arch.mkdir(parents=True)
    (arch / "migration_notes.md").write_text(
        "# 迁移历史对照\n\n> 历史对照：旧路径 src/docwen/gui/core/window.py 已迁移至 "
        "packages/apps/gui/src/docwen_gui/。gui_tk/ 已删除。\n",
        encoding="utf-8",
    )

    result = _run_checker(tmp_path)

    assert result.returncode == 0, (
        f"checker must not flag labeled historical docs; got:\n{result.stdout}\n{result.stderr}"
    )


def test_checker_allows_whitelisted_mapping_doc(tmp_path: Path) -> None:
    """Whitelisted migration-mapping docs may contain old paths in mapping tables."""
    docs = tmp_path / "docs"
    spec = docs / "specs"
    spec.mkdir(parents=True)
    (spec / "资源与打包清单.md").write_text(
        "# 资源与打包清单\n\n| 旧路径 | 新路径 |\n|---|---|\n"
        "| src/docwen/i18n/locales/zh.toml | i18n/locales/zh.toml |\n",
        encoding="utf-8",
    )

    result = _run_checker(tmp_path)

    assert result.returncode == 0, (
        f"checker must not flag whitelisted mapping docs; got:\n{result.stdout}\n{result.stderr}"
    )


def test_checker_allows_archive_dir(tmp_path: Path) -> None:
    """Archived docs (docs/archive/) are exempt regardless of old-path content."""
    docs = tmp_path / "docs"
    archive = docs / "archive"
    archive.mkdir(parents=True)
    (archive / "old_design.md").write_text(
        "# 旧设计（已归档）\n\n引用 src/docwen/gui/core/ 与 theme_styles/。\n",
        encoding="utf-8",
    )

    result = _run_checker(tmp_path)

    assert result.returncode == 0, f"checker must not flag docs/archive/; got:\n{result.stdout}\n{result.stderr}"
