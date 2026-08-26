from __future__ import annotations

import ast
from pathlib import Path

import pytest

pytestmark = pytest.mark.gui


class TestNoOldImports:
    def test_main_window_no_old_imports(self) -> None:
        from docwen_gui import main_window as mw_mod

        source = mw_mod.__file__
        if source is None:
            pytest.skip("Cannot find main_window module source")

        tree = ast.parse(Path(source).read_text(encoding="utf-8"))
        forbidden = {"src.docwen.gui", "docwen.gui.", "docwen.cli."}
        old_mixin = {"mixin", "old_gui", "legacy"}

        for node in ast.walk(tree):
            if isinstance(node, (ast.Import, ast.ImportFrom)):
                module_name = node.module if isinstance(node, ast.ImportFrom) else ""
                for name in node.names if isinstance(node, ast.Import) else []:
                    module_name = name.name
                    for fb in forbidden:
                        assert not module_name.startswith(fb)
                if isinstance(node, ast.ImportFrom) and node.module:
                    for fb in forbidden:
                        assert not node.module.startswith(fb)
                    for om in old_mixin:
                        assert om not in node.module.lower()

    def test_app_module_no_old_imports(self) -> None:
        from docwen_gui import app as app_mod

        source = app_mod.__file__
        if source is None:
            pytest.skip("Cannot find app module source")

        tree = ast.parse(Path(source).read_text(encoding="utf-8"))
        forbidden = {"src.docwen.gui", "docwen.gui.", "docwen.cli."}
        old_mixin = {"mixin", "old_gui", "legacy"}

        for node in ast.walk(tree):
            if isinstance(node, ast.ImportFrom) and node.module:
                for fb in forbidden:
                    assert not node.module.startswith(fb)
                for om in old_mixin:
                    assert om not in node.module.lower()
