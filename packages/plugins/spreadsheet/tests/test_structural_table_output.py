"""Spreadsheet→Structural Tables output regressions."""

from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

from ._xlsx_to_md_golden_support import _build_fake_context

pytestmark = pytest.mark.contract


def _convert_structural(tmp_path: Path, workbook: openpyxl.Workbook) -> str:
    source = tmp_path / "structural.xlsx"
    workbook.save(source)
    workbook.close()
    staging = tmp_path / "stage"
    staging.mkdir()

    from docwen_plugin_spreadsheet.to_markdown.converter import SpreadsheetToMarkdownConverter

    context = _build_fake_context(
        str(source),
        str(staging),
        options={"markdown_extensions": {"output": {"structural_tables": True}}},
        source_format="xlsx",
    )
    result = SpreadsheetToMarkdownConverter().convert(context)

    assert result.success, result.error
    return Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")


def _table_source(markdown: str) -> str:
    rows = [line for line in markdown.splitlines() if line.startswith("|")]
    return "\n".join(rows)


def test_literal_merge_markers_are_escaped_in_structural_output(tmp_path: Path) -> None:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet["A1"] = "Kind"
    sheet["B1"] = "Value"
    sheet["A2"] = "<"
    sheet["B2"] = "^"

    markdown = _convert_structural(tmp_path, workbook)

    assert "| \\< | \\^ |" in markdown
    table = _table_source(markdown)
    assert table.splitlines()[0] == "| Kind | Value |"
    assert table.splitlines()[1] == "| --- | --- |"


def test_merge_crossing_first_row_forces_zero_header_output(tmp_path: Path) -> None:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet["A1"] = "Merged"
    sheet["B1"] = "X"
    sheet["B2"] = "Y"
    sheet.merge_cells("A1:A2")

    markdown = _convert_structural(tmp_path, workbook)
    table = _table_source(markdown)
    lines = table.splitlines()

    assert lines[0] == "| --- | --- |"
    assert lines[1] == "| Merged | X |"
    assert lines[2] == "| ^ | Y |"

