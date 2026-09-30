"""Offline structured clipboard HTML parsing."""

from __future__ import annotations

import pytest

from docwen_core.models.clipboard_document import (
    ClipboardBlock,
    ClipboardDocumentError,
    ClipboardParagraph,
    ClipboardTable,
    ClipboardText,
)
from docwen_gui.clipboard_structured import extract_cf_html_fragment, project_structured_clipboard_html

pytestmark = pytest.mark.contract


def _first_text(block: ClipboardBlock) -> str:
    assert isinstance(block, ClipboardParagraph)
    if not block.inlines:
        return ""
    inline = block.inlines[0]
    assert isinstance(inline, ClipboardText)
    return inline.value


def _cf_html(fragment: str) -> bytes:
    prefix = (
        "Version:1.0\r\n"
        "StartHTML:{start_html:010d}\r\n"
        "EndHTML:{end_html:010d}\r\n"
        "StartFragment:{start_fragment:010d}\r\n"
        "EndFragment:{end_fragment:010d}\r\n"
        "SourceURL:file:///C:/secret/source.docx\r\n"
    )
    html = f"<html><body><!--StartFragment-->{fragment}<!--EndFragment--></body></html>".encode()
    placeholder = prefix.format(start_html=0, end_html=0, start_fragment=0, end_fragment=0).encode()
    start_html = len(placeholder)
    marker_start = html.index(b"<!--StartFragment-->") + len(b"<!--StartFragment-->")
    marker_end = html.index(b"<!--EndFragment-->")
    start_fragment = start_html + marker_start
    end_fragment = start_html + marker_end
    end_html = start_html + len(html)
    header = prefix.format(
        start_html=start_html,
        end_html=end_html,
        start_fragment=start_fragment,
        end_fragment=end_fragment,
    ).encode()
    return header + html


def test_cf_html_uses_byte_offsets_and_does_not_grant_source_url_access() -> None:
    payload = _cf_html("<p>中文</p><table><tr><td>00123</td></tr></table>")
    fragment = extract_cf_html_fragment(payload)
    assert fragment.decode("utf-8").startswith("<p>中文</p>")
    document = project_structured_clipboard_html(payload)
    assert len(document.blocks) == 2
    assert isinstance(document.blocks[0], ClipboardParagraph)
    assert isinstance(document.blocks[1], ClipboardTable)


def test_cf_html_rejects_inconsistent_offsets() -> None:
    payload = _cf_html("<table><tr><td>x</td></tr></table>")
    broken = payload.replace(b"EndFragment:000000", b"EndFragment:999999", 1)
    with pytest.raises(ClipboardDocumentError):
        extract_cf_html_fragment(broken)


def test_rowspan_zero_stops_at_own_row_group_and_nested_order_is_preserved() -> None:
    html = (
        b"<table><tbody>"
        b"<tr><td rowspan='0'><p>before</p><table><tr><td>inner</td></tr></table><p>after</p></td><td>A</td></tr>"
        b"<tr><td>B</td></tr>"
        b"</tbody><tfoot><tr><td>C</td><td>D</td></tr></tfoot></table>"
    )
    document = project_structured_clipboard_html(html)
    outer = document.blocks[0]
    assert isinstance(outer, ClipboardTable)
    assert (outer.row_count, outer.column_count) == (3, 2)
    anchor = outer.cells[0]
    assert anchor.row_span == 2
    assert [type(block) for block in anchor.blocks] == [ClipboardParagraph, ClipboardTable, ClipboardParagraph]
    assert _first_text(anchor.blocks[0]) == "before"
    assert _first_text(anchor.blocks[2]) == "after"


def test_no_header_all_empty_single_row_and_text_fidelity_are_valid() -> None:
    document = project_structured_clipboard_html(
        b"<table><tr><td></td><td> </td><td>&nbsp;</td><td>00123</td><td>$A^2$ | &lt; ^</td></tr></table>"
    )
    table = document.blocks[0]
    assert isinstance(table, ClipboardTable)
    assert (table.row_count, table.column_count) == (1, 5)
    assert table.cells[0].blocks == ()
    assert all(len(cell.blocks) <= 1 for cell in table.cells)
    values = [_first_text(cell.blocks[0]) if cell.blocks else "" for cell in table.cells]
    assert values == ["", " ", "\u00a0", "00123", "$A^2$ | < ^"]


def test_ragged_and_row_group_crossing_tables_fail_closed() -> None:
    with pytest.raises(ClipboardDocumentError, match="incomplete"):
        project_structured_clipboard_html(b"<table><tr><td>A</td><td>B</td></tr><tr><td>C</td></tr></table>")
    with pytest.raises(ClipboardDocumentError, match="row group"):
        project_structured_clipboard_html(
            b"<table><tbody><tr><td rowspan='2'>A</td></tr></tbody><tfoot><tr><td>B</td></tr></tfoot></table>"
        )
