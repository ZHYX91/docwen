"""Guards for the VIS-2026-07-16-092 production warning-clean contracts."""

from __future__ import annotations

import ast
from pathlib import Path
from typing import get_type_hints

import pytest

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "pyright-production-warning-clean-contracts-2026-07-16.md"


def _read(relative_path: str) -> str:
    return (PROJECT_ROOT / relative_path).read_text(encoding="utf-8")


def test_final_production_warning_contracts_stay_explicit() -> None:
    document = _read("packages/plugins/document/src/docwen_plugin_document/to_markdown/converter.py")
    gongwen = _read(
        "packages/plugins/optimizers/gongwen/src/docwen_plugin_optimizer_gongwen/extraction/special_content.py"
    )
    invoice = _read(
        "packages/plugins/optimizers/invoice_cn/src/docwen_plugin_optimizer_invoice_cn/invoice_cn/converter.py"
    )
    gongwen_tests = _read("packages/plugins/optimizers/gongwen/tests/test_gongwen_extraction.py")

    from docwen_core.docx_parsing.list_processing import ListCounterManager, detect_list_item
    from docwen_core.docx_parsing.numbering_index import NumberingLevel

    # Numbering now exposes concrete identities and effective levels. The old
    # abs_ pseudo-ID branch would lose per-instance overrides and counters.
    hints = get_type_hints(NumberingLevel)
    assert hints["num_id"] is str and hints["abstract_num_id"] is str and hints["ilvl"] is int
    assert get_type_hints(detect_list_item)["return"] == tuple[str | None, int | None, str | None]
    assert get_type_hints(ListCounterManager.next)["num_id"] is str
    document_tree = ast.parse(document)
    assert not any(
        isinstance(node, ast.Call)
        and isinstance(node.func, ast.Attribute)
        and node.func.attr == "startswith"
        and node.args
        and isinstance(node.args[0], ast.Constant)
        and node.args[0].value == "abs_"
        for node in ast.walk(document_tree)
    )
    assert any(
        isinstance(node, ast.Call)
        and isinstance(node.func, ast.Attribute)
        and isinstance(node.func.value, ast.Name)
        and node.func.value.id == "numbering_index"
        and node.func.attr == "lookup"
        and node.args
        and isinstance(node.args[0], ast.Name)
        and node.args[0].id == "num_id"
        for node in ast.walk(document_tree)
    )
    assert any(
        isinstance(node, ast.Call)
        and isinstance(node.func, ast.Attribute)
        and isinstance(node.func.value, ast.Name)
        and node.func.value.id == "list_counter"
        and node.func.attr == "next"
        and node.args
        and isinstance(node.args[0], ast.Attribute)
        and isinstance(node.args[0].value, ast.Name)
        and node.args[0].value.id == "level_info"
        and node.args[0].attr == "num_id"
        for node in ast.walk(document_tree)
    )
    native_tests = ast.parse(_read("tests/integration/test_native_docx_reading_contract.py"))
    assert {
        "test_concrete_numbering_overrides_keep_separate_native_counters",
        "test_compound_numbering_uses_the_same_concrete_parent_override",
        "test_numbering_template_empty_missing_literal_and_future_level_have_distinct_meaning",
    } <= {node.name for node in ast.walk(native_tests) if isinstance(node, ast.FunctionDef)}
    assert "isinstance(parent, etree._Element)" not in gongwen
    assert "isinstance(tc, etree._Element)" not in gongwen
    assert "isinstance(tr, etree._Element)" not in gongwen
    assert 'hasattr(parent, "getparent")' not in gongwen
    assert 'hasattr(tc, "getparent")' not in gongwen
    assert 'options.get("yaml_key_labels") if isinstance(options, dict)' not in invoice
    assert 'yaml_key_labels=options.get("yaml_key_labels")' in invoice
    assert "test_repeat_header_row_and_textbox_parent_chain_are_detected" in gongwen_tests
