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


def test_multiple_tables_keep_textual_order_empty_cells_and_values() -> None:
    result = project_clipboard_html(
        "<p>Before</p><table><tr><th>A</th><th>B</th></tr><tr><td></td><td>00123</td></tr></table>"
        "<p>Middle</p><table><tr><th>Pipe</th><th>Lines</th></tr>"
        "<tr><td>x|y</td><td>one<br>two</td></tr></table><p>After</p>"
    )
    assert result.table_count == 2
    assert result.fallback_table_count == 0
    assert result.text.index("Before") < result.text.index("| A | B |") < result.text.index("Middle")
    assert result.text.index("Middle") < result.text.index("| Pipe | Lines |") < result.text.index("After")
    assert "|  | 00123 |" in result.text
    assert "| x\\|y | one<br>two |" in result.text


def test_complex_table_keeps_all_extractable_cell_text() -> None:
    result = project_clipboard_html(
        "<table><tr><th>A</th><th>B</th></tr>"
        "<tr><td rowspan='2'>kept-rowspan</td><td>first</td></tr><tr><td>second</td></tr></table>"
    )
    assert result.fallback_table_count == 1
    assert all(value in result.text for value in ("A", "B", "kept-rowspan", "first", "second"))


def test_script_style_and_remote_images_are_inert_and_reported() -> None:
    result = project_clipboard_html(
        "<p>Before</p><script>DO_NOT_RUN()</script><style>body{display:none}</style>"
        "<img src='https://example.invalid/private.png'><p>After</p>"
    )
    assert result.text == "Before\n\nAfter\n"
    assert "DO_NOT_RUN" not in result.text
    assert "display:none" not in result.text
    assert result.image_count == 1


def test_fallback_table_preserves_empty_spaces_nbsp_pipe_and_line_breaks_exactly() -> None:
    result = project_clipboard_html(
        "<table><tr><td></td><td>  edge  </td><td>&nbsp;x&nbsp;</td>"
        "<td>00123</td><td>x|y</td><td>one<br>two</td></tr></table>"
    )

    assert result.table_count == 0
    assert result.fallback_table_count == 1
    assert result.text == "\t  edge  \t\u00a0x\u00a0\t00123\tx|y\tone\ntwo\n"
