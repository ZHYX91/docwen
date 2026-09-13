"""Ambiguous single-column text uses its validated CSV/TSV declaration."""

from pathlib import Path

import pytest

from docwen_core.detection._sniffing import detect_text_format
from docwen_core.detection._validation import inspect_file

pytestmark = pytest.mark.contract


@pytest.mark.parametrize("suffix", ["csv", "tsv"])
def test_literal_single_column_is_admitted_as_declared_table(tmp_path: Path, suffix: str) -> None:
    source = tmp_path / f"input.{suffix}"
    source.write_text("literal\n1234567890123456789\n00012\n=1+1\n#N/A\n中文\n", encoding="utf-8")
    assert detect_text_format(str(source)) == "txt"
    inspection = inspect_file(str(source))
    assert inspection.detected_format == suffix
    assert inspection.workflow_category == "spreadsheet"
    assert inspection.may_execute and not inspection.requires_explicit_acceptance


def test_quoted_comma_is_a_single_csv_cell(tmp_path: Path) -> None:
    source = tmp_path / "input.csv"
    source.write_text('literal\n"one,two"\n', encoding="utf-8")
    assert inspect_file(str(source)).detected_format == "csv"


@pytest.mark.parametrize("content", ['literal\n"unterminated\n', "# Heading\n\nParagraph\n"])
def test_single_column_declaration_does_not_override_conflicting_content(tmp_path: Path, content: str) -> None:
    source = tmp_path / "input.csv"
    source.write_text(content, encoding="utf-8")
    inspection = inspect_file(str(source))
    assert inspection.detected_format != "csv"
    assert inspection.requires_explicit_acceptance


@pytest.mark.parametrize("text", ["#N/A", "#hashtag", "####### Heading"])
def test_hash_values_are_not_markdown_headings(tmp_path: Path, text: str) -> None:
    source = tmp_path / "input.txt"
    source.write_text(text, encoding="utf-8")
    assert detect_text_format(str(source)) == "txt"
