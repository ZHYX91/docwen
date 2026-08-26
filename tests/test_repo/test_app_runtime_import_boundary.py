"""Governance test: apps layer runtime import boundary.

apps (CLI/GUI) may import from these runtime public API modules ONLY:
- ``docwen_runtime.errors`` — failure categories, security error types
- ``docwen_runtime.security`` — startup security protections
- ``docwen_runtime.config`` — config loading
- ``docwen_runtime.i18n`` — locale management
- ``docwen_runtime.numbering`` — numbering scheme registry
- ``docwen_runtime.templates`` — template registry
- ``docwen_runtime.resources`` — resource resolution
- ``docwen_runtime.logging`` — logging runtime state and path resolution
- ``docwen_runtime.path_io`` — public filesystem syscall path spelling

Imports from other runtime sub-modules (adapters, output_finalizer, etc.)
are forbidden — those are internal implementation details.
"""

from __future__ import annotations

import ast
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

REPO_ROOT = Path(__file__).resolve().parents[2]

# ── Allowed runtime sub-modules for apps ─────────────────────────────

ALLOWED_RUNTIME_SUBS: frozenset[str] = frozenset(
    {
        "docwen_runtime.errors",
        "docwen_runtime.security",
        "docwen_runtime.config",
        "docwen_runtime.i18n",
        "docwen_runtime.numbering",
        "docwen_runtime.templates",
        "docwen_runtime.resources",
        "docwen_runtime.logging",
        "docwen_runtime.path_io",
    }
)

# ── App source roots ─────────────────────────────────────────────────

APP_ROOTS: tuple[str, ...] = (
    "packages/apps/cli/src/docwen_cli",
    "packages/apps/gui/src/docwen_gui",
)


def _collect_runtime_imports(source_dir: Path) -> list[tuple[int, str]]:
    """Return (line_no, import_path) for docwen_runtime imports under *source_dir*."""
    violations: list[tuple[int, str]] = []
    for py_file in sorted(source_dir.rglob("*.py")):
        try:
            tree = ast.parse(py_file.read_text(encoding="utf-8"))
        except SyntaxError:
            continue
        for node in ast.walk(tree):
            if isinstance(node, ast.ImportFrom):
                module = node.module or ""
                if module.startswith("docwen_runtime"):
                    violations.append((node.lineno, module))
            elif isinstance(node, ast.Import):
                for alias in node.names:
                    if alias.name.startswith("docwen_runtime"):
                        violations.append((node.lineno, alias.name))
    return violations


class TestAppRuntimeImportBoundary:
    """apps layer must only import from allowed runtime public API modules."""

    def test_import_collector_catches_import_and_import_from_forms(self, tmp_path: Path) -> None:
        module = tmp_path / "sample.py"
        module.write_text(
            "\n".join(
                [
                    "import docwen_runtime.adapters as adapters",
                    "from docwen_runtime.config import get_config_loader",
                    "import docwen_runtime.logging",
                ]
            ),
            encoding="utf-8",
        )

        assert _collect_runtime_imports(tmp_path) == [
            (1, "docwen_runtime.adapters"),
            (2, "docwen_runtime.config"),
            (3, "docwen_runtime.logging"),
        ]

    def test_cli_only_allowed_runtime_imports(self) -> None:
        """CLI runtime imports must be from the allowed set."""
        cli_root = REPO_ROOT / "packages/apps/cli/src/docwen_cli"
        if not cli_root.is_dir():
            pytest.skip("CLI source not found")
        bad: list[str] = []
        for lineno, module in _collect_runtime_imports(cli_root):
            top = ".".join(module.split(".")[:2])
            if top not in ALLOWED_RUNTIME_SUBS:
                bad.append(f"  line {lineno}: from {module} import ...")
        assert not bad, (
            "docwen_cli imports from disallowed runtime sub-modules:\n"
            + "\n".join(bad)
            + "\n\nAllowed runtime modules: "
            + ", ".join(sorted(ALLOWED_RUNTIME_SUBS))
        )

    def test_gui_only_allowed_runtime_imports(self) -> None:
        """GUI runtime imports must be from the allowed set."""
        gui_root = REPO_ROOT / "packages/apps/gui/src/docwen_gui"
        if not gui_root.is_dir():
            pytest.skip("GUI source not found")
        bad: list[str] = []
        for lineno, module in _collect_runtime_imports(gui_root):
            top = ".".join(module.split(".")[:2])
            if top not in ALLOWED_RUNTIME_SUBS:
                bad.append(f"  line {lineno}: from {module} import ...")
        assert not bad, (
            "docwen_gui imports from disallowed runtime sub-modules:\n"
            + "\n".join(bad)
            + "\n\nAllowed runtime modules: "
            + ", ".join(sorted(ALLOWED_RUNTIME_SUBS))
        )
