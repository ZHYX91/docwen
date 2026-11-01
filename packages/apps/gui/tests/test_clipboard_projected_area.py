"""HTML grid area is reserved before nested content and coverage expansion."""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen_core.models.clipboard_document import ClipboardDocumentError, ClipboardTable
from docwen_gui.clipboard_capture import FrozenClipboardCapture
from docwen_gui.clipboard_rich_document import project_frozen_rich_document
from docwen_gui.clipboard_structured import project_structured_clipboard_html

pytestmark = pytest.mark.contract

OVERSIZED_HTML = b'<table><tr><td rowspan="2" colspan="65536">A</td></tr><tr></tr></table>'


def test_compact_oversized_span_retains_complete_plain_capture(tmp_path: Path) -> None:
    assert len(OVERSIZED_HTML) == 71
    source = tmp_path / "span.html"
    source.write_bytes(OVERSIZED_HTML)
    plain = "  Before\r\nA\r\nAfter\u00a0 "
    plain_source = tmp_path / "plain.txt"
    plain_source.write_bytes(plain.encode("utf-8"))
    with pytest.raises(ClipboardDocumentError) as caught:
        project_structured_clipboard_html(source.read_bytes())
    assert caught.value.code == "clipboard.budget_exceeded"
    capture = FrozenClipboardCapture(plain_source.read_bytes().decode("utf-8"), source.read_bytes(), (), None)
    decision = project_frozen_rich_document(capture)
    assert decision.handled and decision.plain_fallback
    assert decision.projection is None
    assert capture.plain_text == plain
    assert source.read_bytes() == OVERSIZED_HTML


def test_parent_projected_area_is_rejected_before_parsing_an_invalid_child_table(tmp_path: Path) -> None:
    invalid_child = b"<table><tr><td>B</td><td>C</td></tr><tr><td>D</td></tr></table>"
    source = tmp_path / "oversized-with-child.html"
    source.write_bytes(OVERSIZED_HTML.replace(b">A</td>", b">A" + invalid_child + b"</td>"))
    with pytest.raises(ClipboardDocumentError) as caught:
        project_structured_clipboard_html(source.read_bytes())
    # The complete parent shape already exceeds the supported grid. Child
    # semantics must not be evaluated first, even if that child is malformed.
    assert caught.value.code == "clipboard.budget_exceeded"


def test_coverage_area_is_cumulative_and_distinct_from_anchor_count(tmp_path: Path) -> None:
    source = tmp_path / "two-wide-merges.html"
    source.write_bytes(
        b'<table><tr><td colspan="32768">first</td></tr></table><table><tr><td colspan="32769">second</td></tr></table>'
    )
    with pytest.raises(ClipboardDocumentError) as caught:
        project_structured_clipboard_html(source.read_bytes())
    assert caught.value.code == "clipboard.budget_exceeded"


def test_small_merged_cells_keep_coordinates_spans_and_header_evidence(tmp_path: Path) -> None:
    html = (
        b'<table><thead><tr><th colspan="2" id="header" scope="colgroup">Header</th></tr></thead>'
        b'<tbody><tr><th rowspan="0" scope="rowgroup">Group</th><td headers="header">A</td></tr>'
        b'<tr><td headers="header">B</td></tr></tbody><tfoot><tr><td>C</td><td>D</td></tr></tfoot></table>'
    )
    source = tmp_path / "legal.html"
    source.write_bytes(html)
    document = project_structured_clipboard_html(source.read_bytes())
    table = document.blocks[0]
    assert isinstance(table, ClipboardTable)
    assert (table.row_count, table.column_count) == (4, 2)
    assert [(cell.row, cell.column, cell.row_span, cell.column_span) for cell in table.cells] == [
        (0, 0, 1, 2),
        (1, 0, 2, 1),
        (1, 1, 1, 1),
        (2, 1, 1, 1),
        (3, 0, 1, 1),
        (3, 1, 1, 1),
    ]
    assert [(cell.header, cell.scope, cell.headers, cell.html_id) for cell in table.cells] == [
        (True, "colgroup", (), "header"),
        (True, "rowgroup", (), ""),
        (False, "", ("header",), ""),
        (False, "", ("header",), ""),
        (False, "", (), ""),
        (False, "", (), ""),
    ]
