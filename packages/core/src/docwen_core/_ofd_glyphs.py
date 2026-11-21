"""Resolve explicit OFD glyphs from embedded sfnt and standalone CFF subsets."""

from __future__ import annotations

import logging
import math
import struct
from collections.abc import Sequence
from dataclasses import dataclass
from io import BytesIO
from typing import Any

from fontTools.agl import toUnicode
from fontTools.cffLib import CFFFontSet
from fontTools.fontBuilder import FontBuilder
from fontTools.pens.cu2quPen import Cu2QuPen
from fontTools.pens.recordingPen import DecomposingRecordingPen
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.ttLib import TTFont, TTLibError, newTable

logger = logging.getLogger(__name__)

type Outline = tuple[float, Any, str]


def _load_sfnt_header(font: TTFont) -> int:
    # fontTools normalizes old subset timestamps while lazily decoding head
    # and logs to stderr. Apply its same epoch correction in this private font.
    data = font.getTableData("head")
    created, modified = struct.unpack_from(">QQ", data, 20)
    mac_epoch_offset = 0x7C259DC0
    stamps = tuple(
        value + mac_epoch_offset if value < mac_epoch_offset else value
        for value in (created & 0xFFFFFFFF, modified & 0xFFFFFFFF)
    )
    if stamps != (created, modified):
        head = newTable("head")
        head.decompile(data[:20] + struct.pack(">QQ", *stamps) + data[36:], font)
        font["head"] = head
    return int(getattr(font["head"], "unitsPerEm", 0))


@dataclass
class OutlineFont:
    units_per_em: float
    glyph_set: Any
    glyph_order: Sequence[str]
    cmap: dict[int, str]
    cid_names: dict[int, str]
    cubic: bool = False
    ascent: int | None = None
    descent: int | None = None

    @classmethod
    def from_bytes(cls, payload: bytes) -> OutlineFont | None:
        try:
            if payload.startswith(b"\x01\x00"):
                cff = CFFFontSet()
                cff.decompile(BytesIO(payload), None, isCFF2=False)
                top = cff[0]
                xx, xy, yx, yy, dx, dy = top.FontMatrix
                if not (xx == yy and math.isfinite(xx) and xx > 0 and (xy, yx, dx, dy) == (0, 0, 0, 0)):
                    return None
                order = top.charset
                ids = {int(name[3:]): name for name in order if name.startswith("cid") and name[3:].isdigit()}
                cmap = {}
                if not ids:
                    for name in order:
                        value = toUnicode(name)
                        if len(value) == 1:
                            cmap.setdefault(ord(value), name)
                return cls(1 / xx, top.CharStrings, order, cmap, ids, cubic=True)

            font = TTFont(BytesIO(payload), fontNumber=0)
            units = _load_sfnt_header(font)
            if units <= 0:
                return None
            cmap = (font.getBestCmap() or {}) if "cmap" in font else {}
            hhea = font.get("hhea")
            return cls(
                units,
                font.getGlyphSet(),
                font.getGlyphOrder(),
                cmap,
                {},
                "glyf" not in font,
                getattr(hhea, "ascent", None),
                getattr(hhea, "descent", None),
            )
        except (TTLibError, EOFError, ValueError, KeyError, IndexError, TypeError, struct.error, AssertionError):
            logger.debug("Cannot read embedded OFD font outlines", exc_info=True)
            return None

    def glyph_by_unicode(self, codepoint: int) -> Outline | None:
        if name := self.cmap.get(codepoint):
            return self.units_per_em, self.glyph_set, name
        return None

    def glyph_by_id(self, gid: int) -> Outline | None:
        if 0 <= gid < len(self.glyph_order):
            return self.units_per_em, self.glyph_set, self.glyph_order[gid]
        if name := self.cid_names.get(gid):
            return self.units_per_em, self.glyph_set, name
        return None

    def glyphs_by_id(self, gids: Sequence[int]) -> list[Outline | None]:
        # Some producers retain original CIDs when embedding a smaller CFF
        # charset. Resolve the whole run consistently if it uses those CIDs.
        if (
            self.cid_names
            and any(gid >= len(self.glyph_order) for gid in gids)
            and all(gid == 0 or gid in self.cid_names for gid in gids)
        ):
            return [
                (self.units_per_em, self.glyph_set, self.cid_names[gid] if gid else self.glyph_order[0]) for gid in gids
            ]
        return [self.glyph_by_id(gid) for gid in gids]

    def to_truetype(self, mapping: dict[int, str]) -> bytes:
        """Encode the selected source outlines as visible, selectable PDF text."""
        units = round(self.units_per_em)
        builder = FontBuilder(units)
        glyphs = {".notdef": TTGlyphPen(None).glyph()}
        metrics = {".notdef": (units, 0)}
        names = {name: f"g{index}" for index, name in enumerate(dict.fromkeys(mapping.values()), 1)}
        for name, target in names.items():
            source = self.glyph_set[name]
            recording = DecomposingRecordingPen(self.glyph_set)
            source.draw(recording)
            pen = TTGlyphPen(None)
            if self.cubic:
                recording.replay(Cu2QuPen(pen, max_err=units / 1000, reverse_direction=True))
            else:
                recording.replay(pen)
            glyph = pen.glyph()
            glyph.recalcBounds(None)
            glyphs[target] = glyph
            metrics[target] = round(source.width), glyph.xMin

        builder.setupGlyphOrder(list(glyphs))
        builder.setupCharacterMap({codepoint: names[name] for codepoint, name in mapping.items()})
        builder.setupGlyf(glyphs)
        builder.setupHorizontalMetrics(metrics)
        ascent = self.ascent if self.ascent is not None else round(units * 0.8)
        descent = self.descent if self.descent is not None else -round(units * 0.2)
        builder.setupHorizontalHeader(ascent=ascent, descent=descent)
        builder.setupNameTable({"familyName": "OFD", "styleName": "Regular", "psName": "DocWenOFDSubset"})
        builder.setupOS2(
            sTypoAscender=ascent, sTypoDescender=descent, usWinAscent=max(0, ascent), usWinDescent=abs(descent)
        )
        builder.setupPost(keepGlyphNames=False)
        stream = BytesIO()
        builder.save(stream)
        return stream.getvalue()
