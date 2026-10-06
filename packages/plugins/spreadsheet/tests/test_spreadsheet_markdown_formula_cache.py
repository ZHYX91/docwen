"""Spreadsheet→Markdown formula-cache diagnostics."""

from __future__ import annotations

from pathlib import Path

import pytest
from tests.support.formula_cache_fixture import write_formula_cache_fixture

from ._xlsx_to_md_golden_support import _build_fake_context

pytestmark = pytest.mark.contract


def test_xlsx_to_markdown_warns_when_formula_cache_is_unavailable(tmp_path: Path) -> None:
    from docwen_plugin_spreadsheet.to_markdown.converter import SpreadsheetToMarkdownConverter

    source = tmp_path / "cache.xlsx"
    write_formula_cache_fixture(source)
    original = source.read_bytes()
    staging = tmp_path / "stage"
    staging.mkdir()

    result = SpreadsheetToMarkdownConverter().convert(
        _build_fake_context(str(source), str(staging), source_format="xlsx")
    )

    assert result.success, result.error
    warnings = [diagnostic for diagnostic in result.diagnostics if diagnostic.level == "warning"]
    formula_warnings = [
        diagnostic for diagnostic in warnings if diagnostic.code == "SHEET2MD-FORMULA-CACHE-UNAVAILABLE"
    ]
    assert len(formula_warnings) == 1
    message = formula_warnings[0].message
    assert "27 formula cell(s)" in message
    assert "Calc!B1" in message
    assert "Other!A18" in message
    assert "7 more not listed" in message
    assert result.artifacts[0].metadata["formula_cache_unavailable_count"] == 27
    assert result.metrics.extra["formula_cache_unavailable_count"] == 27
    assert source.read_bytes() == original
