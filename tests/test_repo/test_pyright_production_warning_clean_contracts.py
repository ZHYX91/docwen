"""Guards for the VIS-2026-07-16-092 production warning-clean contracts."""

from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "pyright-production-warning-clean-contracts-2026-07-16.md"


def _read(relative_path: str) -> str:
    return (PROJECT_ROOT / relative_path).read_text(encoding="utf-8")


def test_final_production_warning_contracts_stay_explicit() -> None:
    document = _read("packages/plugins/document/src/docwen_plugin_document/to_markdown/converter.py")
    markdown = _read("packages/plugins/markdown/src/docwen_plugin_markdown/template_filler.py")
    gongwen = _read(
        "packages/plugins/optimizers/gongwen/src/docwen_plugin_optimizer_gongwen/extraction/special_content.py"
    )
    invoice = _read(
        "packages/plugins/optimizers/invoice_cn/src/docwen_plugin_optimizer_invoice_cn/invoice_cn/converter.py"
    )
    markdown_tests = _read("packages/plugins/markdown/tests/test_field_registry.py")
    gongwen_tests = _read("packages/plugins/optimizers/gongwen/tests/test_gongwen_extraction.py")

    assert "and isinstance(num_id, str)" not in document
    assert 'num_id.startswith("abs_")' in document
    assert "not isinstance(key, str)" not in markdown
    assert "if not callable(handler):" in markdown
    assert "isinstance(parent, etree._Element)" not in gongwen
    assert "isinstance(tc, etree._Element)" not in gongwen
    assert "isinstance(tr, etree._Element)" not in gongwen
    assert 'hasattr(parent, "getparent")' not in gongwen
    assert 'hasattr(tc, "getparent")' not in gongwen
    assert 'options.get("yaml_key_labels") if isinstance(options, dict)' not in invoice
    assert 'yaml_key_labels=options.get("yaml_key_labels")' in invoice
    assert "test_special_placeholder_dispatch_keeps_non_callable_handler_guard" in markdown_tests
    assert "test_repeat_header_row_and_textbox_parent_chain_are_detected" in gongwen_tests
