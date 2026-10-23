"""Excel text capacity must never result in a successful truncated export."""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from .conftest import make_context, write_temp_md

pytestmark = pytest.mark.contract


@pytest.mark.parametrize(
    "content",
    ["x" * 32767 + "IMPORTANT-TAIL", "😀" * 16384],
)
def test_markdown_table_cell_over_capacity_fails_with_location(content: str) -> None:
    from docwen_plugin_markdown.to_spreadsheet.converter import MdToXlsxConverter

    source = write_temp_md(f"| Value |\n| --- |\n| {content} |\n")
    context, _workspace = make_context(source, target_format="xlsx")
    result = MdToXlsxConverter().convert(context)

    assert not result.success
    assert result.error is not None
    assert result.error.diagnostic_code == "MD2XLSX-CELL-TEXT-TOO-LONG"
    assert "row 2, column 1" in result.error.message
    assert "32767" in result.error.message
    assert not result.artifacts


def test_exactly_maximum_ascii_cell_is_preserved() -> None:
    from docwen_plugin_markdown.to_spreadsheet.converter import MdToXlsxConverter

    value = "x" * 32767
    source = write_temp_md(f"| Value |\n| --- |\n| {value} |\n")
    context, _workspace = make_context(source, target_format="xlsx")
    result = MdToXlsxConverter().convert(context)
    assert result.success, result.error
    workbook = load_workbook(result.artifacts[0].staging_path)
    try:
        assert workbook.active["A2"].value == value
    finally:
        workbook.close()


def test_template_column_fill_over_capacity_fails_before_openpyxl_truncates(tmp_path: Path) -> None:
    from docwen_plugin_markdown.to_spreadsheet.converter import MdToXlsxConverter

    template = tmp_path / "template.xlsx"
    workbook = Workbook()
    workbook.active["A1"] = "{{↓Value}}"
    workbook.save(template)
    workbook.close()

    source = write_temp_md(f"| Value |\n| --- |\n| {'x' * 32768} |\n")
    context, _workspace = make_context(
        source, target_format="xlsx", options={"template_name": str(template)}
    )
    result = MdToXlsxConverter().convert(context)
    assert not result.success
    assert result.error is not None
    assert result.error.diagnostic_code == "MD2XLSX-CELL-TEXT-TOO-LONG"
    assert not result.artifacts
