"""Guard safe exclusion of unused build-time/Tk runtimes from frozen bundles."""

from __future__ import annotations

import ast
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

ROOT = Path(__file__).resolve().parents[2]
EXCLUDED = {"tkinter", "turtle", "idlelib", "Cython"}


def test_frozen_builder_excludes_build_compiler_and_legacy_tk_toolkit() -> None:
    source = (ROOT / "scripts/build/build.py").read_text(encoding="utf-8")
    for name in EXCLUDED:
        assert source.count(f'"--exclude-module={name}"') == 1
    # The same exclusions must apply to both independently analyzed entries.
    assert source.count("*common_excludes,") == 2


def test_product_source_does_not_import_excluded_modules() -> None:
    imported: list[str] = []
    for path in (ROOT / "packages").glob("**/src/**/*.py"):
        parsed = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(parsed):
            if isinstance(node, ast.Import):
                imports = [item.name for item in node.names]
            elif isinstance(node, ast.ImportFrom):
                imports = [node.module or ""]
            else:
                continue
            for name in imports:
                if name.split(".")[0] in EXCLUDED:
                    imported.append(f"{path.relative_to(ROOT)}:{node.lineno}:{name}")
    assert imported == []
