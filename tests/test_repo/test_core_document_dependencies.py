"""Core's shared DOCX facade must work without installing a format plugin."""

from __future__ import annotations

import ast
import subprocess
import sys
import textwrap
import tomllib
from pathlib import Path

import pytest
from packaging.requirements import Requirement

pytestmark = pytest.mark.unit

ROOT = Path(__file__).resolve().parents[2]
CORE = ROOT / "packages/core"


def test_core_declares_the_document_libraries_it_imports_directly() -> None:
    metadata = tomllib.loads((CORE / "pyproject.toml").read_text(encoding="utf-8"))
    declared = {Requirement(value).name.casefold() for value in metadata["project"]["dependencies"]}
    document_libraries = {"docx": "python-docx", "lxml": "lxml"}
    missing: set[str] = set()
    for path in (CORE / "src/docwen_core").rglob("*.py"):
        for node in ast.walk(ast.parse(path.read_text(encoding="utf-8"))):
            imports = (
                [alias.name for alias in node.names]
                if isinstance(node, ast.Import)
                else [node.module or ""]
                if isinstance(node, ast.ImportFrom) and node.level == 0
                else []
            )
            for imported in imports:
                distribution = document_libraries.get(imported.split(".", 1)[0])
                if distribution is not None and distribution not in declared:
                    missing.add(distribution)
    assert not missing, f"Core must declare its directly imported document libraries: {sorted(missing)}"


def test_core_docx_facade_imports_and_projects_styles_without_other_product_packages(tmp_path: Path) -> None:
    # -I excludes PYTHONPATH, the current directory and user-site customizations.
    # Explicitly block other product packages even in a workspace-wide dev install.
    probe = textwrap.dedent("""\
        import importlib.abc
        import pathlib
        import sys

        class CoreOnly(importlib.abc.MetaPathFinder):
            def find_spec(self, fullname, path=None, target=None):
                if fullname.startswith("docwen_") and fullname.split(".")[0] != "docwen_core":
                    raise ModuleNotFoundError("Other product packages are unavailable: " + fullname)
                return None

        sys.meta_path.insert(0, CoreOnly())
        source_root = pathlib.Path(sys.argv[1]).resolve()
        sys.path.insert(0, str(source_root))
        import docwen_core.docx_semantics_v3 as semantics
        from docx import Document
        from lxml import etree
        from docwen_core._docx_semantics_v3_heading_styles import (
            capture_heading_style_binding,
            heading_style_binding_map_xml,
            parse_heading_style_binding_map,
        )

        assert pathlib.Path(semantics.__file__).resolve().is_relative_to(source_root)
        document = Document()
        paragraph = document.add_heading("Shared semantic heading", level=1)
        binding = capture_heading_style_binding(document, paragraph, "heading-1", 1)
        assert isinstance(binding, semantics.HeadingStyleBindingV3)
        encoded = heading_style_binding_map_xml((binding,))
        assert parse_heading_style_binding_map(etree.fromstring(encoded)) == (binding,)
        assert not any(name.startswith("docwen_") and name.split(".")[0] != "docwen_core"
                       for name in sys.modules)
        """)
    completed = subprocess.run(
        [sys.executable, "-I", "-c", probe, str(CORE / "src")],
        cwd=tmp_path,
        capture_output=True,
        text=True,
        timeout=30,
        check=False,
    )
    assert completed.returncode == 0, completed.stdout + completed.stderr
