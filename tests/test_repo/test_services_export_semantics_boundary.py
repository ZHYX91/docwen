"""治理测试：workspace packages 不得依赖旧 monolith 模块。"""

from __future__ import annotations

import ast
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

_FORBIDDEN_PREFIXES = (
    "docwen.application",
    "docwen.bootstrap",
    "docwen.cli",
    "docwen.config",
    "docwen.converter",
    "docwen.core",
    "docwen.docx_spell",
    "docwen.errors",
    "docwen.formats",
    "docwen.gui",
    "docwen.i18n",
    "docwen.ipc",
    "docwen.md_spell",
    "docwen.security",
    "docwen.services",
    "docwen.template",
    "docwen.text_rules",
    "docwen.utils",
)


def _collect_workspace_source_files() -> list[Path]:
    project_root = Path(__file__).resolve().parents[2]
    packages_dir = project_root / "packages"
    return sorted(packages_dir.rglob("src/**/*.py"))


def _is_forbidden(module: str | None) -> bool:
    if not module:
        return False
    return module.startswith(_FORBIDDEN_PREFIXES)


def test_workspace_packages_do_not_import_legacy_monolith_modules() -> None:
    project_root = Path(__file__).resolve().parents[2]
    violations: list[str] = []
    for path in _collect_workspace_source_files():
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(tree):
            if isinstance(node, ast.ImportFrom):
                if _is_forbidden(node.module):
                    violations.append(f"{path.relative_to(project_root)}:{node.lineno} from {node.module}")
            elif isinstance(node, ast.Import):
                for alias in node.names:
                    if _is_forbidden(alias.name):
                        violations.append(f"{path.relative_to(project_root)}:{node.lineno} import {alias.name}")

    assert not violations, "workspace packages must not import legacy monolith modules:\n" + "\n".join(violations)
