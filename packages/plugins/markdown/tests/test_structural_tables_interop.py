from __future__ import annotations

import pytest

from docwen_plugin_markdown.document_semantics import analyze_document_semantics
from docwen_plugin_markdown.mistune_extensions import parse_markdown_text

pytestmark = pytest.mark.unit


def _table(source: str):
    ast = parse_markdown_text(source)
    assert len(ast) == 1
    return ast[0]


def test_structural_table_dialect_projects_multi_row_and_row_headers() -> None:
    source = "| Region | Sales | < |\n| Quarter | Q1 | Q2 |\n| --- || --- | --- |\n| North | 10 | 12 |\n| ^ | 8 | 11 |"
    analysis = analyze_document_semantics(parse_markdown_text(source), current_v3=True)

    assert not analysis.has_errors
    metadata = analysis.ast[0]["_document_semantics_table"]
    assert metadata["header_rows"] == 2
    assert metadata["header_columns"] == 1
    assert metadata["row_count"] == 4
    assert metadata["column_count"] == 3
    assert metadata["anchors"] == [
        {
            "row": 0,
            "column": 0,
            "row_span": 1,
            "column_span": 1,
            "role": "corner_header",
            "children": [{"type": "text", "raw": "Region"}],
        },
        {
            "row": 0,
            "column": 1,
            "row_span": 1,
            "column_span": 2,
            "role": "column_header",
            "children": [{"type": "text", "raw": "Sales"}],
        },
        {
            "row": 1,
            "column": 0,
            "row_span": 1,
            "column_span": 1,
            "role": "corner_header",
            "children": [{"type": "text", "raw": "Quarter"}],
        },
        {
            "row": 1,
            "column": 1,
            "row_span": 1,
            "column_span": 1,
            "role": "column_header",
            "children": [{"type": "text", "raw": "Q1"}],
        },
        {
            "row": 1,
            "column": 2,
            "row_span": 1,
            "column_span": 1,
            "role": "column_header",
            "children": [{"type": "text", "raw": "Q2"}],
        },
        {
            "row": 2,
            "column": 0,
            "row_span": 2,
            "column_span": 1,
            "role": "row_header",
            "children": [{"type": "text", "raw": "North"}],
        },
        {
            "row": 2,
            "column": 1,
            "row_span": 1,
            "column_span": 1,
            "role": "data",
            "children": [{"type": "text", "raw": "10"}],
        },
        {
            "row": 2,
            "column": 2,
            "row_span": 1,
            "column_span": 1,
            "role": "data",
            "children": [{"type": "text", "raw": "12"}],
        },
        {
            "row": 3,
            "column": 1,
            "row_span": 1,
            "column_span": 1,
            "role": "data",
            "children": [{"type": "text", "raw": "8"}],
        },
        {
            "row": 3,
            "column": 2,
            "row_span": 1,
            "column_span": 1,
            "role": "data",
            "children": [{"type": "text", "raw": "11"}],
        },
    ]


def test_structural_table_escaped_merge_markers_remain_literal() -> None:
    source = "| Group | \\< |\n| Name | Value |\n| --- | --- |\n| A | \\^ |"
    analysis = analyze_document_semantics(parse_markdown_text(source), current_v3=True)

    assert not analysis.has_errors
    metadata = analysis.ast[0]["_document_semantics_table"]
    assert all(anchor["row_span"] == anchor["column_span"] == 1 for anchor in metadata["anchors"])
    assert [
        child["raw"] for anchor in metadata["anchors"] for child in anchor["children"] if child.get("type") == "text"
    ] == ["Group", "<", "Name", "Value", "A", "^"]


def test_ordinary_and_invalid_structural_tables_remain_outside_the_extension() -> None:
    ordinary = _table("| A | B |\n| --- | --- |\n| 1 | 2 |")
    invalid = _table("| A | B |\n| --- || --- |\n| 1 | 2 | 3 |")

    assert ordinary["type"] == "table"
    assert "_structural_table" not in ordinary
    assert invalid["type"] == "paragraph"
