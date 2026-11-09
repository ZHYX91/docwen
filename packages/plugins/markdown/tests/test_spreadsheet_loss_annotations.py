"""Spreadsheet losses retain data, existing notes and occurrence diagnostics."""

from __future__ import annotations

import csv
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.worksheet.worksheet import Worksheet

from docwen_plugin_markdown.to_spreadsheet.converter import MdToCsvConverter, MdToXlsxConverter
from docwen_plugin_markdown.to_spreadsheet.losses import SpreadsheetLosses

from .conftest import make_context

pytestmark = pytest.mark.contract


def active(workbook: Workbook) -> Worksheet:
    worksheet = workbook.active
    assert isinstance(worksheet, Worksheet)
    return worksheet


def convert(tmp_path: Path, *, target="xlsx", comments=True, template=None, content=None):
    source = tmp_path / "source.md"
    source.write_text(content or "| Formula |\n| --- |\n| $x^2$ |\n", encoding="utf-8")
    context, _ = make_context(
        str(source),
        target_format=target,
        config_values={"conversion": {"loss_comments": comments}},
        options={"template_name": str(template)} if template else {},
    )
    result = (MdToXlsxConverter() if target == "xlsx" else MdToCsvConverter()).convert(context)
    assert result.success, result.error
    return result


@pytest.mark.parametrize("comments", [True, False])
def test_formula_text_loss_is_reported_independently_of_comments(tmp_path, comments):
    result = convert(tmp_path, comments=comments)
    cell = load_workbook(result.artifacts[0].staging_path)["Sheet1"]["A2"]
    assert cell.value == "$x^2$"
    assert (cell.comment is not None) == comments
    losses = [d for d in result.diagnostics if d.code == "CONVERSION-LOSS-SPREADSHEET-MATH"]
    assert len(losses) == 1
    assert losses[0].location == "Sheet1!A2"


def test_csv_diagnostics_do_not_add_rows_or_columns(tmp_path):
    result = convert(tmp_path, target="csv")
    with open(result.artifacts[0].staging_path, encoding="utf-8-sig", newline="") as stream:
        assert list(csv.reader(stream)) == [["Formula"], ["$x^2$"]]
    assert any(d.code == "CONVERSION-LOSS-SPREADSHEET-MATH" for d in result.diagnostics)


def test_literal_code_and_user_link_mode_do_not_warn(tmp_path):
    result = convert(tmp_path, content="| Value |\n| --- |\n| `$x^2$` |\n| plain text |\n")
    assert not any(d.code.startswith("CONVERSION-LOSS") for d in result.diagnostics)


def test_comment_is_appended_and_same_loss_is_not_duplicated():
    cell = active(Workbook())["A1"]
    cell.comment = Comment("My review", "Reviewer")
    losses = SpreadsheetLosses(comments=True)
    losses.formulas("$x^2$", "Sheet!A1", cell)
    losses.formulas("$x^2$", "Sheet!A1", cell)
    assert cell.comment.text.startswith("My review\n\nDocWen:")
    assert cell.comment.author == "Reviewer"
    assert len(losses.collector.records) == 1


@pytest.mark.parametrize("comments", [True, False])
def test_protected_template_loss_preserves_original_comment(tmp_path, comments):
    wb = Workbook()
    sheet = active(wb)
    sheet["A1"] = "{{↓Formula}}"
    sheet["A1"].comment = Comment("Original review", "Reviewer")
    sheet.protection.sheet = True
    template = tmp_path / "template.xlsx"
    wb.save(template)
    result = convert(tmp_path, template=template, comments=comments)
    cell = active(load_workbook(result.artifacts[0].staging_path))["A1"]
    assert cell.comment.text.startswith("Original review")
    assert cell.comment.author == "Reviewer"
    assert ("DocWen:" in cell.comment.text) == comments
    assert any(d.code == "CONVERSION-LOSS-PROTECTED-CELL" for d in result.diagnostics)


def test_comment_write_failure_keeps_loss_diagnostic():
    class FailingCell:
        @property
        def comment(self):
            return None

        @comment.setter
        def comment(self, value):
            raise OSError("cannot write comment")

    losses = SpreadsheetLosses(comments=True)
    losses.formulas("$x^2$", "Sheet!A1", FailingCell())
    assert len(losses.collector.records) == 1
    assert losses.annotated_count == 0
    assert {d.code for d in losses.diagnostics} == {
        "CONVERSION-LOSS-SPREADSHEET-MATH",
        "CONVERSION-LOSS-COMMENT-FAILED",
        "CONVERSION-LOSS-SUMMARY",
    }


def test_untouched_template_math_does_not_get_conversion_comment(tmp_path):
    wb = Workbook()
    active(wb)["A1"] = "$existing$"
    active(wb)["B1"] = "{{name}}"
    template = tmp_path / "template.xlsx"
    wb.save(template)
    result = convert(tmp_path, template=template, content="---\nname: Plain\n---\n")
    assert active(load_workbook(result.artifacts[0].staging_path))["A1"].comment is None
    assert not any(d.code.startswith("CONVERSION-LOSS") for d in result.diagnostics)


def test_csv_template_merge_loss_has_location_and_keeps_geometry(tmp_path):
    wb = Workbook()
    sheet = active(wb)
    sheet["A1"] = "Merged"
    sheet.merge_cells("A1:B1")
    template = tmp_path / "template.xlsx"
    wb.save(template)
    result = convert(tmp_path, target="csv", template=template)
    loss = next(d for d in result.diagnostics if d.code == "CONVERSION-LOSS-CSV-MERGE")
    assert loss.location.endswith("!A1:B1")
    with open(result.artifacts[0].staging_path, encoding="utf-8-sig", newline="") as stream:
        assert list(csv.reader(stream)) == [["Merged", ""]]


@pytest.mark.parametrize("target", ["xlsx", "csv"])
@pytest.mark.parametrize("mode", ["extract_text", "keep"])
def test_selected_link_text_is_not_a_formula_render_loss(tmp_path, target, mode):
    source = tmp_path / "source.md"
    source.write_text("| Value |\n| --- |\n| [Energy $mc^2$](https://example.com) and $x$ |\n", encoding="utf-8")
    context, _ = make_context(
        str(source), target_format=target, config_values={"link": {"non_embed_links": {"markdown_mode": mode}}}
    )
    result = (MdToXlsxConverter() if target == "xlsx" else MdToCsvConverter()).convert(context)
    assert result.success
    losses = [d for d in result.diagnostics if d.code == "CONVERSION-LOSS-SPREADSHEET-MATH"]
    assert len(losses) == 1
    assert "mc^2" not in losses[0].message
    assert "Original: x" in losses[0].message


def test_invalid_structural_merge_has_cell_annotation(tmp_path):
    source = tmp_path / "source.md"
    source.write_text("| A | B |\n| --- | --- |\n| < | value |\n", encoding="utf-8")
    context, _ = make_context(
        str(source), target_format="xlsx", options={"markdown_extensions": {"input": {"structural_tables": True}}}
    )
    result = MdToXlsxConverter().convert(context)
    assert result.success
    loss = next(d for d in result.diagnostics if d.code == "CONVERSION-LOSS-SPREADSHEET-MERGE")
    assert loss.location == "Sheet1!A2"
    cell = load_workbook(result.artifacts[0].staging_path)["Sheet1"]["A2"]
    assert cell.value == "<"
    assert cell.comment is not None


def test_template_merge_conflict_has_retained_anchor_annotation(tmp_path):
    wb = Workbook()
    ws = active(wb)
    ws["A1"] = "{{↓A}}"
    ws["B1"] = "{{↓B}}"
    ws.merge_cells("A2:A3")
    ws["A2"] = "keep"
    template = tmp_path / "template.xlsx"
    wb.save(template)
    source = tmp_path / "source.md"
    source.write_text("| A | B |\n| --- | --- |\n| one | two |\n| merge | < |\n", encoding="utf-8")
    context, _ = make_context(
        str(source),
        target_format="xlsx",
        options={"template_name": str(template), "markdown_extensions": {"input": {"structural_tables": True}}},
    )
    result = MdToXlsxConverter().convert(context)
    assert result.success, result.error
    assert any(d.code == "CONVERSION-LOSS-SPREADSHEET-MERGE" for d in result.diagnostics)
    cell = active(load_workbook(result.artifacts[0].staging_path))["A2"]
    assert cell.comment is not None
