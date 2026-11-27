"""Built-in plugins may use DocWen Core and their own implementation."""

from __future__ import annotations

import ast
import os
import sys
from pathlib import Path
from textwrap import dedent

import pytest
from tests.support.subprocess_runner import run_subprocess

from docwen_bundle.runtime_factory import _DEFAULT_PLUGIN_IMPORTS

pytestmark = pytest.mark.contract

REPO_ROOT = Path(__file__).resolve().parents[2]
PLUGIN_ROOTS = {root.name: root for root in (REPO_ROOT / "packages/plugins").glob("**/src/docwen_plugin_*")}


@pytest.mark.parametrize("plugin", _DEFAULT_PLUGIN_IMPORTS)
def test_plugin_source_uses_only_core_and_local_modules(plugin: str) -> None:
    violations = []
    root = PLUGIN_ROOTS[plugin]
    for path in sorted(root.rglob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                modules = [alias.name for alias in node.names]
            elif isinstance(node, ast.ImportFrom) and not node.level and node.module:
                modules = [node.module]
            else:
                continue
            for module in modules:
                package = module.partition(".")[0]
                if (package == "docwen" or package.startswith("docwen_")) and package not in {"docwen_core", plugin}:
                    violations.append(f"{path.relative_to(REPO_ROOT)}:{node.lineno}: {module}")
    assert not violations, "\n".join(violations)


@pytest.mark.parametrize("module", [*_DEFAULT_PLUGIN_IMPORTS, "docwen_plugin_document.to_document.converter"])
def test_plugin_import_in_a_clean_interpreter(module: str) -> None:
    plugin = module.partition(".")[0]
    probe = dedent(
        """
        import importlib
        import sys
        from pathlib import Path

        module, expected_root = sys.argv[1:]
        plugin = module.partition(".")[0]
        loaded_module = importlib.import_module(module)
        if not Path(loaded_module.__file__).resolve().is_relative_to(expected_root):
            raise RuntimeError(f"{module} resolved outside the checkout: {loaded_module.__file__}")
        local = {name.partition(".")[0] for name in sys.modules if name == "docwen" or name.startswith("docwen_")}
        forbidden = local - {plugin, "docwen_core"}
        if forbidden:
            raise RuntimeError(f"{module} imported forbidden packages: {sorted(forbidden)}")
        """
    )
    root = PLUGIN_ROOTS[plugin]
    result = run_subprocess(
        [sys.executable, "-c", probe, module, str(root)],
        env={**os.environ, "PYTHONPATH": os.pathsep.join([str(root.parent), *sys.path])},
    )
    assert result.returncode == 0, result.stderr
