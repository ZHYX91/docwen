"""Governance test: every real package test directory is covered by pytest testpaths.

Prevents the testpaths from rotting — stale entries pointing to
non-existent directories, or real test directories silently excluded
from the default ``pytest`` run.
"""

from __future__ import annotations

from pathlib import Path

import pytest
import tomlkit

pytestmark = pytest.mark.contract

REPO_ROOT = Path(__file__).resolve().parents[2]
PYPROJECT = REPO_ROOT / "pyproject.toml"


def _load_testpaths() -> list[str]:
    """Return the raw list of testpaths from pyproject.toml."""
    data = tomlkit.parse(PYPROJECT.read_text(encoding="utf-8"))
    return list(data["tool"]["pytest"]["ini_options"]["testpaths"])


def _existing_test_dirs() -> list[Path]:
    """Discover all real test directories under packages/."""
    test_dirs: list[Path] = []
    packages = REPO_ROOT / "packages"
    if not packages.is_dir():
        return test_dirs

    for test_dir in sorted(packages.rglob("tests")):
        if test_dir.is_dir() and any(test_dir.glob("test_*.py")):
            test_dirs.append(test_dir)
    for test_dir in sorted((REPO_ROOT / "tests").rglob("*")):
        if test_dir.is_dir() and any(test_dir.glob("test_*.py")):
            test_dirs.append(REPO_ROOT / "tests")
            break
    return test_dirs


def _plugin_test_dirs_importing_repo_test_support() -> list[Path]:
    """Return plugin test dirs that import shared ``tests.support`` helpers."""
    plugin_root = REPO_ROOT / "packages" / "plugins"
    dirs: list[Path] = []
    for test_dir in sorted(plugin_root.glob("*/tests")):
        if not test_dir.is_dir():
            continue
        for path in test_dir.rglob("*.py"):
            text = path.read_text(encoding="utf-8", errors="ignore")
            if "tests.support" in text:
                dirs.append(test_dir)
                break
    return dirs


class TestPytestTestpathCoverage:
    """Ensure testpaths in pyproject.toml match reality."""

    def test_no_stale_testpath_entries(self) -> None:
        """Every entry in testpaths must point to an existing directory."""
        testpaths = _load_testpaths()
        stale: list[str] = []
        for tp in testpaths:
            p = REPO_ROOT / tp
            if not p.is_dir():
                stale.append(tp)
        assert not stale, (
            f"pyproject.toml [tool.pytest.ini_options] testpaths contains "
            f"{len(stale)} stale entr{'y' if len(stale) == 1 else 'ies'} "
            f"pointing to non-existent director{'y' if len(stale) == 1 else 'ies'}:\n"
            + "\n".join(f"  - {s}" for s in stale)
            + "\n\nRemove or rename the entries above."
        )

    def test_all_plugin_test_dirs_covered(self) -> None:
        """Every packages/plugins/*/tests/ dir with test files is in testpaths.

        There are no implicit exclusions: an existing plugin test directory must
        participate in the default suite.
        """
        testpaths_set = {tp.replace("\\", "/") for tp in _load_testpaths()}
        uncovered: list[str] = []
        for d in _existing_test_dirs():
            rel = str(d.relative_to(REPO_ROOT).as_posix())
            if rel not in testpaths_set:
                uncovered.append(rel)
        assert not uncovered, (
            f"{len(uncovered)} test director{'y' if len(uncovered) == 1 else 'ies'} "
            f"exist but are NOT in pyproject.toml testpaths:\n"
            + "\n".join(f"  + {u}" for u in uncovered)
            + "\n\nAdd the missing directories to [tool.pytest.ini_options] testpaths "
            "so they are included in the default ``pytest`` run."
        )

    def test_plugin_tests_using_shared_support_seed_repo_root(self) -> None:
        """Plugin-focused pytest runs must be able to import ``tests.support``.

        Several plugin packages carry their own ``tool.pytest.ini_options`` and
        can be run directly.  When a plugin test imports shared fakes from the
        repository ``tests.support`` package, its local conftest must seed the
        repository root just like the document/spreadsheet/layout plugin tests.
        """
        missing: list[str] = []
        for test_dir in _plugin_test_dirs_importing_repo_test_support():
            conftest = test_dir / "conftest.py"
            if not conftest.exists():
                missing.append(f"{test_dir.relative_to(REPO_ROOT).as_posix()}: missing conftest.py")
                continue
            text = conftest.read_text(encoding="utf-8")
            required_snippets = (
                "PROJECT_ROOT = Path(__file__).resolve().parents[4]",
                "LOCAL_SRC_PATHS",
                "PROJECT_ROOT,",
            )
            if not all(snippet in text for snippet in required_snippets):
                missing.append(f"{test_dir.relative_to(REPO_ROOT).as_posix()}: conftest does not seed repo root")

        assert not missing, "\n".join(missing)
