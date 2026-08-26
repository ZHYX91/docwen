"""Tests for list processing: ListCounterManager, and detect_list_item including
pStyle fallback.
"""

import re
from unittest.mock import MagicMock

import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from docwen_core.text.heading_numbering import strip_heading_prefix
from docwen_plugin_document.shared.list_processing import (
    ListCounterManager,
    detect_list_item,
    format_list_marker,
)
from docwen_plugin_document.shared.numbering_index import (
    NumberingLevel,
)

pytestmark = pytest.mark.unit

# ── ListCounterManager ───────────────────────────────────────────────────


def test_list_counter_resets_deeper_levels_when_parent_increments():
    counter = ListCounterManager()
    assert counter.next("1", 0) == 1
    assert counter.next("1", 1) == 1
    assert counter.next("1", 0) == 2
    assert counter.next("1", 1) == 1  # reset because parent incremented


def test_list_counter_different_num_ids():
    counter = ListCounterManager()
    assert counter.next("A", 0) == 1
    assert counter.next("B", 0) == 1
    assert counter.next("A", 0) == 2


def test_list_counter_reset_clears_all():
    counter = ListCounterManager()
    counter.next("1", 0)
    counter.next("2", 1)
    counter.reset()
    assert counter.next("1", 0) == 1
    assert counter.next("2", 1) == 1


def test_list_counter_honors_level_start_and_then_increments():
    counter = ListCounterManager()
    assert counter.next("custom", 0, start=3) == 3
    assert counter.next("custom", 0, start=99) == 4


def test_list_counter_peek_and_snapshot_do_not_mutate_state():
    counter = ListCounterManager()
    assert counter.peek("custom", 1, start=4) == 4
    assert counter.snapshot("custom") == {}
    assert counter.next("custom", 1, start=4) == 4
    assert counter.peek("custom", 1, start=99) == 5
    snapshot = counter.snapshot("custom")
    assert snapshot == {1: 4}
    snapshot[1] = 500
    assert counter.snapshot("custom") == {1: 4}


def test_format_list_marker_accepts_configured_unordered_styles():
    assert format_list_marker("bullet", 0, "dash") == "-"
    assert format_list_marker("bullet", 0, "asterisk") == "*"
    assert format_list_marker("bullet", 0, "plus") == "+"
    assert format_list_marker("bullet", 0, "unknown") == "-"
    assert format_list_marker("ordered", 3, "plus") == "3."


_HEADING_RULES = (
    ("chinese", re.compile(r"^[一二三四五六七八九十]+、"), 1),
    ("digit", re.compile(r"^\d+\."), 1),
    ("circled", re.compile(r"^[①②③④⑤⑥⑦⑧⑨⑩]"), 1),
)


def test_strip_heading_numbering_chinese():
    assert strip_heading_prefix("一、总体要求", rules=_HEADING_RULES) == ("一、", "总体要求")


def test_strip_heading_numbering_digit():
    # Note: core's strip_heading_prefix does not strip whitespace
    # from the remaining text, so the space after "1." is preserved.
    assert strip_heading_prefix("1. 概述", rules=_HEADING_RULES) == ("1.", " 概述")


def test_strip_heading_numbering_circled():
    assert strip_heading_prefix("①标题", rules=_HEADING_RULES) == ("①", "标题")


def test_strip_heading_numbering_no_match():
    assert strip_heading_prefix("plain text", rules=_HEADING_RULES) == ("", "plain text")


# ── detect_list_item pStyle fallback ─────────────────────────────────────


def _make_mock_para_without_numpr(style_id: str = "1Heading1"):
    """Build a MagicMock Paragraph where pPr exists but numPr does not,
    and para.style has the given *style_id*."""
    para = MagicMock()

    # para._p.find("{w}pPr") returns a mock pPr
    pPr_mock = MagicMock()
    pPr_mock.find.return_value = None  # numPr not found → triggers fallback
    para._p.find.return_value = pPr_mock

    # para.style has style_id
    mock_style = MagicMock()
    mock_style.style_id = style_id
    para.style = mock_style

    return para


def test_detect_list_item_pstyle_fallback_returns_abs_num_id():
    """When a paragraph has no numPr but its style links to numbering via
    pStyle, detect_list_item falls back to lookup_by_style_id and returns
    a synthetic ``abs_{abstract_num_id}`` numId."""
    para = _make_mock_para_without_numpr("1Heading1")

    num_idx = MagicMock()
    level_info = NumberingLevel(
        num_id="20",
        abstract_num_id="1",
        ilvl=0,
        num_fmt="chineseCountingThousand",
        lvl_text="%1、",
    )
    num_idx.lookup_by_style_id.return_value = level_info

    num_id, ilvl, list_type = detect_list_item(para, numbering_index=num_idx)
    assert num_id == "abs_1"
    assert ilvl == 0
    assert list_type == "ordered"


def test_detect_list_item_pstyle_fallback_bullet():
    """pStyle fallback returns 'bullet' when num_fmt is 'bullet'."""
    para = _make_mock_para_without_numpr("ListBullet")

    num_idx = MagicMock()
    level_info = NumberingLevel(
        num_id="5",
        abstract_num_id="2",
        ilvl=0,
        num_fmt="bullet",
        lvl_text="•",
    )
    num_idx.lookup_by_style_id.return_value = level_info

    num_id, ilvl, list_type = detect_list_item(para, numbering_index=num_idx)
    assert num_id == "abs_2"
    assert ilvl == 0
    assert list_type == "bullet"


def test_detect_list_item_pstyle_fallback_no_match():
    """When style_id has no matching pStyle in numbering, returns (None, None, None)."""
    para = _make_mock_para_without_numpr("NormalStyle")

    num_idx = MagicMock()
    num_idx.lookup_by_style_id.return_value = None

    num_id, ilvl, list_type = detect_list_item(para, numbering_index=num_idx)
    assert num_id is None
    assert ilvl is None
    assert list_type is None


def test_detect_list_item_pstyle_fallback_no_numbering_index():
    """Without a numbering_index, no pStyle fallback occurs."""
    para = _make_mock_para_without_numpr("SomeStyle")

    num_id, ilvl, list_type = detect_list_item(para, numbering_index=None)
    assert num_id is None
    assert ilvl is None
    assert list_type is None


def test_detect_list_item_treats_num_id_zero_as_numbering_disabled():
    """OOXML numId=0 cancels numbering and must not become a fallback bullet."""
    doc = Document()
    para = doc.add_paragraph("ordinary body after a list")
    num_pr = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl")
    ilvl.set(qn("w:val"), "0")
    num_id = OxmlElement("w:numId")
    num_id.set(qn("w:val"), "0")
    num_pr.extend((ilvl, num_id))
    para._p.get_or_add_pPr().append(num_pr)
    numbering_index = MagicMock()

    assert detect_list_item(para, numbering_index) == (None, None, None)
    numbering_index.lookup.assert_not_called()
