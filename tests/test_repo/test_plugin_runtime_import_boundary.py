"""Governance test: plugin src layer runtime import boundary.

plugin src may NOT import from any docwen_runtime sub-module.
runtime is a peer layer, not a dependency of plugins.
Static system schemes / rule data must come via docwen_core or
runtime-time injection, never via direct runtime import.
"""

from __future__ import annotations

import ast
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

REPO_ROOT = Path(__file__).resolve().parents[2]

PLUGIN_SRC_ROOTS = [
    "packages/plugins/document/src/docwen_plugin_document",
    "packages/plugins/spreadsheet/src/docwen_plugin_spreadsheet",
    "packages/plugins/image/src/docwen_plugin_image",
    "packages/plugins/print/src/docwen_plugin_print",
    "packages/plugins/layout/src/docwen_plugin_layout",
    "packages/plugins/markdown/src/docwen_plugin_markdown",
    "packages/plugins/markup/src/docwen_plugin_markup",
    "packages/plugins/presentation/src/docwen_plugin_presentation",
    "packages/plugins/proofread/src/docwen_plugin_proofread",
    "packages/plugins/optimizers/gongwen/src/docwen_plugin_optimizer_gongwen",
    "packages/plugins/optimizers/invoice_cn/src/docwen_plugin_optimizer_invoice_cn",
]


def _collect_runtime_imports(source_dir: Path) -> list[tuple[int, str]]:
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


class TestPluginRuntimeImportBoundary:
    """plugin src must not import from any docwen_runtime sub-module."""

    @pytest.mark.parametrize("src_root", PLUGIN_SRC_ROOTS)
    def test_plugin_src_no_runtime_imports(self, src_root: str) -> None:
        root = REPO_ROOT / src_root
        if not root.is_dir():
            pytest.skip(f"{src_root} not found")
        bad = _collect_runtime_imports(root)
        assert not bad, f"{src_root} imports from docwen_runtime (forbidden):\n" + "\n".join(
            f"  line {ln}: {mod}" for ln, mod in bad
        )
