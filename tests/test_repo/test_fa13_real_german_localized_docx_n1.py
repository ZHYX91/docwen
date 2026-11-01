"""Fail-closed evidence guards for VIS-180 / finite-contract FA-13-N1."""

from __future__ import annotations

import ast
from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = [pytest.mark.contract, pytest.mark.golden]

PROJECT_ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "fa13-real-german-localized-docx-n1-current-fix-2026-07-23.md"
STAGE_CARD = "fa13-real-german-localized-docx-n1-stage-card-2026-07-23.md"
STATUS = "N1_FIXED_AND_VERIFIED_N2_PENDING_EXTERNAL_SAMPLE_PERMISSION"
FINAL_STATUS = "FIXED_AND_VERIFIED_WITH_USER_ACCEPTED_DERIVED_IMAGE_ACCURACY_BOUNDARY"
SOURCE_SHA = "44A4CB870A0D3E58EF2B1B719543A5899186F543DC92148287B1C357F719310A"
OUTPUT_SHA = "52190B6A9160F88573F4B5FACCFBC454283026F598C379C275AEAED4258F992D"
SUMMARY_SHA = "C0A8161DB47B7646EA677337FEF60F83E34BB2833528065226A4C0471CAD742B"


def _read(path: Path) -> str:
    return read_source_text(path)


def _function(tree: ast.AST, name: str) -> ast.FunctionDef:
    [function] = [node for node in ast.walk(tree) if isinstance(node, ast.FunctionDef) and node.name == name]
    return function


def test_fa13_n1_has_direct_mixed_revision_regression() -> None:
    renderer = _read(
        PROJECT_ROOT
        / "packages"
        / "plugins"
        / "document"
        / "src"
        / "docwen_plugin_document"
        / "shared"
        / "markdown_runs.py"
    )
    converter = _read(
        PROJECT_ROOT
        / "packages"
        / "plugins"
        / "document"
        / "src"
        / "docwen_plugin_document"
        / "to_markdown"
        / "converter.py"
    )
    tests = _read(
        PROJECT_ROOT / "packages" / "plugins" / "document" / "tests" / "test_to_markdown_standard_parity_*.py"
    )

    assert "def _flush_text()" in renderer
    assert 'elif tag == "tab"' in renderer
    # All paragraphs now use one accepted-visible XML projection, including
    # an empty result. Do not restore the former local wrapper whitelist or a
    # raw paragraph.text fallback merely to satisfy a source spelling check.
    converter_tree = ast.parse(converter)
    process = _function(converter_tree, "_process_paragraph")
    [visible] = [
        node
        for node in ast.walk(process)
        if isinstance(node, ast.Assign)
        and any(isinstance(target, ast.Name) and target.id == "visible_paragraph_text" for target in node.targets)
    ]
    assert isinstance(visible.value, ast.Call) and isinstance(visible.value.func, ast.Attribute)
    assert visible.value.func.attr == "_extract_paragraph_text_formatted"
    assert any(
        isinstance(node, ast.Assign)
        and any(isinstance(target, ast.Name) and target.id == "paragraph_text" for target in node.targets)
        and isinstance(node.value, ast.Call)
        and isinstance(node.value.func, ast.Attribute)
        and isinstance(node.value.func.value, ast.Name)
        and node.value.func.value.id == "visible_paragraph_text"
        and node.value.func.attr == "strip"
        and node.lineno > visible.lineno
        for node in ast.walk(process)
    )
    delegate = _function(converter_tree, "_extract_paragraph_text_formatted")
    [rendered] = [node for node in ast.walk(delegate) if isinstance(node, ast.Return)]
    assert isinstance(rendered.value, ast.Call) and isinstance(rendered.value.func, ast.Name)
    assert rendered.value.func.id == "render_paragraph_runs"
    children = _function(ast.parse(renderer), "_process_children")
    accepted_branches = [
        node
        for node in ast.walk(children)
        if isinstance(node, ast.If)
        and isinstance(node.test, ast.Compare)
        and isinstance(node.test.left, ast.Name)
        and node.test.left.id == "tag"
        and len(node.test.ops) == 1
        and isinstance(node.test.ops[0], ast.In)
        and len(node.test.comparators) == 1
        and isinstance(node.test.comparators[0], ast.Tuple)
        and {"ins", "moveTo", "sdt", "sdtContent", "customXml"}
        <= {value.value for value in node.test.comparators[0].elts if isinstance(value, ast.Constant)}
    ]
    [accepted] = accepted_branches
    assert any(
        isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id == "_process_children"
        for statement in accepted.body
        for node in ast.walk(statement)
    )
    for tag in ("fldSimple", "hyperlink"):
        assert any(
            isinstance(node, ast.Compare)
            and isinstance(node.left, ast.Name)
            and node.left.id == "tag"
            and len(node.ops) == 1
            and isinstance(node.ops[0], ast.Eq)
            and len(node.comparators) == 1
            and isinstance(node.comparators[0], ast.Constant)
            and node.comparators[0].value == tag
            for node in ast.walk(children)
        )
    localized = ast.parse(
        _read(PROJECT_ROOT / "packages/plugins/document/tests/test_to_markdown_standard_parity_02.py")
    )
    _function(localized, "test_docx_to_md_convert_uses_every_bundled_locale_template_style")
    assert "test_mixed_direct_and_inserted_revision_text_reaches_paragraph_processing" in tests
    assert 'assert "Rejected heading" not in' in tests
