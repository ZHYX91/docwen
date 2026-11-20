from io import BytesIO

import pytest
from fontTools.fontBuilder import FontBuilder
from fontTools.pens.recordingPen import RecordingPen
from fontTools.pens.t2CharStringPen import T2CharStringPen
from fontTools.pens.ttGlyphPen import TTGlyphPen

from docwen_core._ofd_glyphs import OutlineFont

pytestmark = pytest.mark.unit


def _triangle(pen):
    pen.moveTo((0, 0))
    pen.lineTo((200, 0))
    pen.lineTo((0, 300))
    pen.closePath()


def _assert_triangle(outline):
    assert outline is not None
    units, glyphs, name = outline
    assert units == 1000
    recorded = RecordingPen()
    glyphs[name].draw(recorded)
    assert ("lineTo", ((200, 0),)) in recorded.value


def test_truetype_subset_without_cmap_uses_declared_glyph_index(caplog):
    builder = FontBuilder(1000)
    builder.setupGlyphOrder([".notdef", "triangle"])
    pen = TTGlyphPen(None)
    _triangle(pen)
    builder.setupGlyf({".notdef": TTGlyphPen(None).glyph(), "triangle": pen.glyph()})
    builder.setupHorizontalMetrics({".notdef": (500, 0), "triangle": (500, 0)})
    builder.setupHorizontalHeader(ascent=800, descent=-200)
    builder.font.recalcTimestamp = False
    builder.font["head"].created = builder.font["head"].modified = 0
    buffer = BytesIO()
    builder.save(buffer)

    font = OutlineFont.from_bytes(buffer.getvalue())
    assert font is not None
    assert not any("timestamp" in record.getMessage() for record in caplog.records)
    _assert_triangle(font.glyph_by_id(1))
    assert font.glyph_by_id(-1) is None
    assert font.glyph_by_id(2) is None
    assert font.glyph_by_unicode(ord("A")) is None
    mapped = OutlineFont.from_bytes(font.to_truetype({ord("文"): font.glyph_order[1]}))
    assert mapped is not None
    _assert_triangle(mapped.glyph_by_unicode(ord("文")))


def test_standalone_cff_draws_declared_glyph_and_resolves_adobe_name():
    builder = FontBuilder(1000, isTTF=False)
    builder.setupGlyphOrder([".notdef", "A"])
    pen = T2CharStringPen(500, None)
    _triangle(pen)
    builder.setupCFF(
        "Synthetic", {}, {".notdef": T2CharStringPen(500, None).getCharString(), "A": pen.getCharString()}, {}
    )
    font = OutlineFont.from_bytes(builder.font.getTableData("CFF "))
    assert font is not None
    _assert_triangle(font.glyph_by_id(1))
    _assert_triangle(font.glyph_by_unicode(ord("A")))
    mapped = OutlineFont.from_bytes(font.to_truetype({ord("文"): "A"}))
    assert mapped is not None
    _assert_triangle(mapped.glyph_by_unicode(ord("文")))


def test_sparse_cid_run_uses_one_numbering_scheme_for_every_glyph():
    order = [".notdef", "cid00010", "cid00001"]
    subsets = OutlineFont(1000, {}, order, {}, {10: order[1], 1: order[2]})
    # 1 can mean either glyph slot 1 or CID 1. The neighbouring 10 resolves it.
    assert [glyph[-1] for glyph in subsets.glyphs_by_id([1, 10, 0]) if glyph] == [order[2], order[1], order[0]]
    assert [glyph[-1] for glyph in subsets.glyphs_by_id([1, 2]) if glyph] == order[1:]
