"""Fail-closed evidence guards for VIS-175 / finite-contract FA-02."""

from __future__ import annotations

import ast
from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = pytest.mark.contract

PROJECT_ROOT = Path(__file__).resolve().parents[2]


def _read(path: Path) -> str:
    return read_source_text(path)


def _function(tree: ast.AST, name: str) -> ast.FunctionDef:
    [function] = [node for node in ast.walk(tree) if isinstance(node, ast.FunctionDef) and node.name == name]
    return function


def test_fa02_accepted_view_repairs_have_direct_guards() -> None:
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
    textbox = _read(
        PROJECT_ROOT / "packages" / "core" / "src" / "docwen_core" / "docx_parsing" / "textbox_extraction.py"
    )
    breaks = _read(PROJECT_ROOT / "packages" / "core" / "src" / "docwen_core" / "docx_parsing" / "break_utils.py")
    parity_tests = _read(
        PROJECT_ROOT / "packages" / "plugins" / "document" / "tests" / "test_to_markdown_standard_parity_*.py"
    )
    textbox_tests = _read(PROJECT_ROOT / "packages" / "core" / "tests" / "test_docx_textbox_extraction.py")

    assert 'elif tag in ("ins", "moveTo", "smartTag", "sdt", "sdtContent", "customXml"):' in renderer
    assert 'elif tag == "fldSimple":' in renderer
    assert 'resolver = getattr(note_extractor, "get_noteref_text", None)' in renderer
    assert "else:\n                    _process_children(child)" in renderer
    assert 'elif tag == "noBreakHyphen":' in renderer
    assert "return vanish is not None and _on_off_property_is_enabled(vanish)" in renderer
    assert "if _run_is_hidden(run):\n            return" in renderer
    # The accepted view keeps authored XML occurrences. Text equality and a
    # shared paragraph anchor are not authority to discard another paragraph.
    textbox_tree = ast.parse(textbox)
    owner = _function(textbox_tree, "_extract_from_part")
    [seen] = [
        node
        for node in ast.walk(owner)
        if isinstance(node, ast.AnnAssign)
        and isinstance(node.target, ast.Name)
        and node.target.id == "processed_elements"
    ]
    assert ast.dump(seen.annotation) == ast.dump(ast.parse("set[object]", mode="eval").body)
    occurrence = _function(textbox_tree, "_extract_from_txbx")
    [membership] = [
        node
        for node in ast.walk(occurrence)
        if isinstance(node, ast.Compare)
        and isinstance(node.left, ast.Name)
        and node.left.id == "para_elem"
        and len(node.ops) == 1
        and isinstance(node.ops[0], ast.In)
        and len(node.comparators) == 1
        and isinstance(node.comparators[0], ast.Name)
        and node.comparators[0].id == "processed_elements"
    ]
    assert membership is not None
    [remember] = [
        node
        for node in ast.walk(occurrence)
        if isinstance(node, ast.Call)
        and isinstance(node.func, ast.Attribute)
        and isinstance(node.func.value, ast.Name)
        and node.func.value.id == "processed_elements"
        and node.func.attr == "add"
    ]
    assert len(remember.args) == 1
    assert isinstance(remember.args[0], ast.Name) and remember.args[0].id == "para_elem"
    # The complete carrier regression supplies the A/B/A, shared-anchor and
    # Choice/Fallback content oracles; this source projection is not a host run.
    native_tests = ast.parse(_read(PROJECT_ROOT / "tests/integration/test_native_textbox_carriers.py"))
    _function(native_tests, "test_complete_native_textbox_carriers_preserve_occurrence_and_document_order")
    assert 'if local_name in {"del", "moveFrom"}:' in textbox
    assert 'if local_name in {"del", "moveFrom"}:' in breaks
    assert "test_deleted_page_break_does_not_erase_accepted_inserted_paragraph" in parity_tests
    assert "test_drawingml_textbox_uses_accepted_revision_and_visible_run_projection" in textbox_tests


def test_fa02_resource_repair_has_direct_executable_guards() -> None:
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
    parity_tests = _read(
        PROJECT_ROOT / "packages" / "plugins" / "document" / "tests" / "test_to_markdown_standard_parity_*.py"
    )
    golden_tests = _read(
        PROJECT_ROOT / "packages" / "plugins" / "document" / "tests" / "test_document_to_md_golden_*.py"
    )

    assert "for index in range(len(lines) - 1, -1, -1):" in converter
    assert "merged[index] += merged_refs" in converter
    assert "return [merged_refs, *lines]" in converter
    assert "test_image_refs_merge_once_after_last_nonempty_rendered_line" in parity_tests
    assert "test_image_only_element_still_emits_reachable_reference_line" in parity_tests
    assert "test_image_only_docx_paragraph_is_finalized_and_linked" in golden_tests
