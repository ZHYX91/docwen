"""Clipboard HTML projection keeps reliable tables without guessing structure."""

from __future__ import annotations

import pytest

from docwen_gui.clipboard_content import project_clipboard_html

pytestmark = pytest.mark.unit


def test_preserves_text_table_text_order_for_explicit_header_table() -> None:
    result = project_clipboard_html(
        "<p>Before</p><table><tr><th>Name</th><th>Value</th></tr><tr><td>A</td><td>00123</td></tr></table><p>After</p>"
    )

    assert result.table_count == 1
    assert result.fallback_table_count == 0
    assert result.text == ("Before\n\n| Name | Value |\n| --- | --- |\n| A | 00123 |\n\nAfter\n")


def test_table_without_explicit_header_falls_back_to_plain_cell_grid() -> None:
    result = project_clipboard_html("<table><tr><td>A</td><td>B</td></tr><tr><td>1</td><td>2</td></tr></table>")

    assert result.table_count == 0
    assert result.fallback_table_count == 1
    assert result.text == "A\tB\n1\t2\n"


@pytest.mark.parametrize(
    "html_text",
    [
        "<table><tr><th>A</th><th>B</th></tr><tr><td colspan='2'>wide</td></tr></table>",
        "<table><tr><th>A</th><th>B</th></tr><tr><td rowspan='2'>tall</td><td>x</td></tr><tr><td>y</td></tr></table>",
        "<table><tr><th>A</th><th>B</th></tr><tr><td>outer<table><tr><td>inner</td></tr></table></td><td>x</td></tr></table>",
    ],
)
def test_complex_tables_fall_back_instead_of_inventing_markdown_structure(html_text: str) -> None:
    result = project_clipboard_html(html_text)

    assert result.table_count == 0
    assert result.fallback_table_count == 1
    assert "|" not in result.text


def test_markdown_table_cells_escape_pipe_and_keep_line_breaks() -> None:
    result = project_clipboard_html(
        "<table><tr><th>A|B</th><th>Lines</th></tr><tr><td>x|y</td><td>one<br>two</td></tr></table>"
    )

    assert result.table_count == 1
    assert "| A\\|B | Lines |" in result.text
    assert "| x\\|y | one<br>two |" in result.text
