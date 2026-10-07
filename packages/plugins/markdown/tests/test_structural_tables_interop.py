from __future__ import annotations

import pytest

from docwen_plugin_markdown.common_utils import parse_raw_md_tables
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


def test_structural_table_dialect_accepts_no_column_header_rows() -> None:
    source = "| - | - |\n| Alice | 10 |\n| Bob | 20 |"
    analysis = analyze_document_semantics(parse_markdown_text(source), current_v3=True)

    assert not analysis.has_errors
    metadata = analysis.ast[0]["_document_semantics_table"]
    assert metadata["header_rows"] == 0
    assert metadata["header_columns"] == 0
    assert all(anchor["role"] == "data" for anchor in metadata["anchors"])


def test_structural_table_accepts_short_delimiters_without_outer_pipes() -> None:
    source = "Region | Sales | <\nQuarter | Q1 | Q2\n- | - | -\nNorth | 10 | 12"
    analysis = analyze_document_semantics(parse_markdown_text(source), current_v3=True)

    assert not analysis.has_errors
    metadata = analysis.ast[0]["_document_semantics_table"]
    assert metadata["header_rows"] == 2
    assert metadata["header_columns"] == 0
    assert metadata["anchors"][0]["column_span"] == 1
    assert metadata["anchors"][1]["column_span"] == 2


def test_formatted_merge_markers_are_literal_cell_content() -> None:
    source = """| Code left | Strong left | Link left | Math left | Code up | Strong up |
| --- | --- | --- | --- | --- | --- |
| `<` | **<** | [<](https://example.com) | $<$ | `^` | **^** |"""
    analysis = analyze_document_semantics(parse_markdown_text(source), current_v3=True)

    assert not analysis.has_errors
    metadata = analysis.ast[0]["_document_semantics_table"]
    for anchor in metadata["anchors"]:
        assert anchor["row_span"] == anchor["column_span"] == 1


def _nested_tables(nodes):
    found = []
    for node in nodes:
        if node.get("type") == "table":
            found.append(node)
        found.extend(_nested_tables(node.get("children", [])))
    return found


@pytest.mark.parametrize(
    "source",
    [
        "> | - | - |\n> | A | B |\n> | C | D |",
        "> [!note]\n>\n> | - | - |\n> | A | B |\n> | C | D |",
        "- Item\n\n  | - | - |\n  | A | B |\n  | C | D |",
    ],
)
def test_structural_tables_inside_quote_callout_and_list_are_annotated(source: str) -> None:
    analysis = analyze_document_semantics(parse_markdown_text(source), current_v3=True)

    assert not analysis.has_errors
    tables = _nested_tables(analysis.ast)
    assert len(tables) == 1
    metadata = tables[0]["_document_semantics_table"]
    assert metadata["header_rows"] == 0
    assert metadata["column_count"] == 2
    assert metadata["row_count"] == 2


@pytest.mark.parametrize(
    ("source", "expected_kind", "expected_raw"),
    [
        ("| - | - |\n| `literal | value |\n| left | right |", "text", "`literal"),
        ("| - | - |\n| `C:\\` | value |\n| left | right |", "codespan", "C:\\"),
    ],
)
def test_structural_table_backtick_boundaries_do_not_consume_column_pipes(
    source: str,
    expected_kind: str,
    expected_raw: str,
) -> None:
    analysis = analyze_document_semantics(parse_markdown_text(source), current_v3=True)

    assert not analysis.has_errors
    metadata = analysis.ast[0]["_document_semantics_table"]
    assert metadata["column_count"] == 2
    first_anchor = next(anchor for anchor in metadata["anchors"] if anchor["row"] == 0 and anchor["column"] == 0)
    assert [(child.get("type"), child.get("raw")) for child in first_anchor["children"]] == [
        (expected_kind, expected_raw)
    ]


def test_escaped_backtick_does_not_open_a_structural_cell_code_span() -> None:
    source = "| - | - |\n| \\`plain | literal` |\n| left | right |"
    analysis = analyze_document_semantics(parse_markdown_text(source), current_v3=True)

    assert not analysis.has_errors
    [table] = _nested_tables(analysis.ast)
    assert table["_document_semantics_table"]["column_count"] == 2


@pytest.mark.parametrize("prefix", ["", "> "])
def test_comment_tables_remain_visible_literal_source(prefix: str) -> None:
    source = "\n".join(prefix + line for line in ["%%", "| A | < |", "| - | - |", "| 1 | 2 |", "%%"])
    ast = parse_markdown_text(source)

    assert not _nested_tables(ast)

    def text(nodes):
        return "".join(node.get("raw", "") + text(node.get("children", [])) for node in nodes)

    assert "| A | < |" in text(ast)
    assert "%%" in text(ast)


@pytest.mark.parametrize(
    "protected",
    [
        "%%\n```\n| - | - |\n| hidden | value |\n%%\n",
        "```md\n%%\n| - | - |\n| hidden | value |\n%%\n```\n",
    ],
)
def test_comment_and_fence_boundaries_do_not_hide_a_later_table(protected: str) -> None:
    source = protected + "\n| - | - |\n| visible | value |\n"
    analysis = analyze_document_semantics(parse_markdown_text(source), current_v3=True)

    assert not analysis.has_errors
    assert len(_nested_tables(analysis.ast)) == 1
    [table] = parse_raw_md_tables(source, structural_tables=True)
    assert table["all_rows"] == [["visible", "value"]]
