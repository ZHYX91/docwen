"""Word numbering.xml index for numPr → list-type/level resolution."""

from __future__ import annotations

from dataclasses import dataclass

from docwen_core.text.numbering import (
    number_to_chinese,
    number_to_circled,
    number_to_letter_lower,
    number_to_letter_upper,
    number_to_roman_lower,
    number_to_roman_upper,
)

# ── Numbering index ─────────────────────────────────────────────────────


@dataclass(frozen=True)
class NumberingLevel:
    num_id: str
    abstract_num_id: str
    ilvl: int
    num_fmt: str
    lvl_text: str
    start: int = 1
    suff: str = "tab"
    p_style: str | None = None


class NumberingIndex:
    """Parse ``word/numbering.xml`` to resolve numId → abstractNumId → level."""

    def __init__(self, doc) -> None:
        self._num_to_abstract: dict[str, str] = {}  # numId → abstractNumId
        self._abstract_levels: dict[str, dict[int, dict]] = {}  # abstractNumId → {ilvl: lvl_elem_attrs}
        self._abstract_num_style_links: dict[str, str] = {}  # abstractNumId → numStyleLink pStyle

        try:
            numbering_part = doc.part.numbering_part
            if numbering_part is None:
                return
            numbering_xml = numbering_part._element
        except (AttributeError, KeyError, NotImplementedError):
            return

        # num → abstractNum
        for num_elem in numbering_xml.findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}num"):
            num_id = num_elem.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numId")
            abs_ref = num_elem.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}abstractNumId")
            if num_id is not None and abs_ref is not None:
                abs_id = abs_ref.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")
                if abs_id is not None:
                    self._num_to_abstract[str(num_id)] = str(abs_id)

        # abstractNum levels
        for abs_elem in numbering_xml.findall(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}abstractNum"
        ):
            abs_id = abs_elem.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}abstractNumId")
            if abs_id is None:
                continue

            levels: dict[int, dict] = {}
            for lvl_elem in abs_elem.findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lvl"):
                ilvl_raw = lvl_elem.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ilvl")
                if ilvl_raw is None:
                    continue
                ilvl = int(ilvl_raw)
                num_fmt = ""
                lvl_text = ""
                start = 1
                suff = "tab"
                p_style = ""
                fmt_elem = lvl_elem.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numFmt")
                if fmt_elem is not None:
                    num_fmt = fmt_elem.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") or ""
                lt_elem = lvl_elem.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lvlText")
                if lt_elem is not None:
                    lvl_text = lt_elem.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") or ""
                start_elem = lvl_elem.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}start")
                if start_elem is not None:
                    start_raw = start_elem.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")
                    if start_raw:
                        try:
                            start = int(start_raw)
                        except ValueError:
                            start = 1
                suff_elem = lvl_elem.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}suff")
                if suff_elem is not None:
                    suff = suff_elem.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") or "tab"
                ps_elem = lvl_elem.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle")
                if ps_elem is not None:
                    p_style = ps_elem.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") or ""
                levels[ilvl] = {
                    "numFmt": num_fmt,
                    "lvlText": lvl_text,
                    "start": start,
                    "suff": suff,
                    "pStyle": p_style,
                }

            self._abstract_levels[str(abs_id)] = levels

            # numStyleLink
            ns_link = abs_elem.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numStyleLink")
            if ns_link is not None:
                self._abstract_num_style_links[str(abs_id)] = (
                    ns_link.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") or ""
                )

    def lookup(self, num_id: str, ilvl: int) -> NumberingLevel | None:
        abs_id = self._num_to_abstract.get(str(num_id))
        if abs_id is None:
            return None
        levels = self._abstract_levels.get(abs_id, {})
        lvl = levels.get(ilvl)
        if lvl is None:
            return None
        return NumberingLevel(
            num_id=str(num_id),
            abstract_num_id=abs_id,
            ilvl=ilvl,
            num_fmt=lvl.get("numFmt", ""),
            lvl_text=lvl.get("lvlText", ""),
            start=lvl.get("start", 1),
            suff=lvl.get("suff", "tab"),
            p_style=lvl.get("pStyle") or None,
        )

    def lookup_by_style_id(self, style_id: str) -> NumberingLevel | None:
        for abs_id, levels in self._abstract_levels.items():
            for ilvl, lvl in levels.items():
                if lvl.get("pStyle") == style_id:
                    # Find a numId that references this abstractNumId
                    for num_id, aid in self._num_to_abstract.items():
                        if aid == abs_id:
                            return NumberingLevel(
                                num_id=num_id,
                                abstract_num_id=abs_id,
                                ilvl=ilvl,
                                num_fmt=lvl.get("numFmt", ""),
                                lvl_text=lvl.get("lvlText", ""),
                                start=lvl.get("start", 1),
                                suff=lvl.get("suff", "tab"),
                                p_style=style_id,
                            )
        return None

    def lookup_by_num_id(self, num_id: str, ilvl: int) -> NumberingLevel | None:
        """Direct lookup by numId and ilvl (convenience alias for ``lookup``)."""
        return self.lookup(num_id, ilvl)

    def preview_numbering_text(self, level: NumberingLevel, ordinal: int = 1) -> str:
        fmt = level.num_fmt.lower() if level.num_fmt else "decimal"
        if fmt == "chineseCountingThousand":
            return number_to_chinese(ordinal)
        elif fmt == "decimal":
            return str(ordinal)
        elif fmt in ("upperLetter", "lowerLetter"):
            if fmt == "upperLetter":
                return number_to_letter_upper(ordinal)
            return number_to_letter_lower(ordinal)
        elif fmt in ("upperRoman", "lowerRoman"):
            if fmt == "upperRoman":
                return number_to_roman_upper(ordinal)
            return number_to_roman_lower(ordinal)
        return str(ordinal)

    def render_numbering_text(
        self,
        level: NumberingLevel,
        counter_value: int,
        parent_counters: dict[int, int] | None = None,
    ) -> str:
        """Resolve ``lvlText`` placeholders against counter values.

        Substitutes ``%1``..``%9`` placeholders with formatted sibling-level
        counters, falling back to *counter_value* for the current level.

        Args:
            level: The ``NumberingLevel`` to render.
            counter_value: The counter value for *level.ilvl*.
            parent_counters: Optional ``{ilvl: value}`` map for sibling levels.

        Returns:
            The resolved numbering prefix text (e.g. ``"1.1 "``, ``"一、"``).
        """
        lvl_text = level.lvl_text
        if not lvl_text:
            return _apply_numbering_suffix(self.preview_numbering_text(level, counter_value), level.suff)

        parents = parent_counters or {}
        result = lvl_text
        for placeholder_idx in range(1, 10):
            ph = f"%{placeholder_idx}"
            if ph not in result:
                continue
            sibling_ilvl = placeholder_idx - 1
            val = counter_value if sibling_ilvl == level.ilvl else parents.get(sibling_ilvl, 0)
            sibling = self.lookup_by_abstract(level.abstract_num_id, sibling_ilvl)
            if val == 0:
                val = sibling.start if sibling else 1
            fmt = sibling.num_fmt if sibling else "decimal"
            formatted = _format_counter(val, fmt)
            result = result.replace(ph, formatted)
        return _apply_numbering_suffix(result, level.suff)

    def lookup_by_abstract(self, abstract_num_id: str, ilvl: int) -> NumberingLevel | None:
        """Look up a numbering level by abstractNumId and ilvl."""
        levels = self._abstract_levels.get(abstract_num_id, {})
        lvl = levels.get(ilvl)
        if lvl is None:
            return None
        num_id = ""
        for nid, aid in self._num_to_abstract.items():
            if aid == abstract_num_id:
                num_id = nid
                break
        return NumberingLevel(
            num_id=num_id,
            abstract_num_id=abstract_num_id,
            ilvl=ilvl,
            num_fmt=lvl.get("numFmt", ""),
            lvl_text=lvl.get("lvlText", ""),
            start=lvl.get("start", 1),
            suff=lvl.get("suff", "tab"),
            p_style=lvl.get("pStyle") or None,
        )


def _apply_numbering_suffix(text: str, suff: str) -> str:
    """Project Word's tab/space suffix to one Markdown-safe ASCII space."""
    return f"{text} " if suff.casefold() in {"tab", "space"} else text


def _format_counter(n: int, num_fmt: str) -> str:
    """Format a counter value according to the Word numFmt code."""
    fmt = (num_fmt or "").lower()
    if fmt in ("chinesecountingthousand", "chinesecounting", "ideographtraditional"):
        return number_to_chinese(n)
    if fmt == "decimal":
        return str(n)
    if fmt == "lowerletter":
        return number_to_letter_lower(n)
    if fmt == "upperletter":
        return number_to_letter_upper(n)
    if fmt == "lowerroman":
        return number_to_roman_lower(n)
    if fmt == "upperroman":
        return number_to_roman_upper(n)
    if fmt == "decimalenclosedcircle":
        return number_to_circled(n)
    return str(n)
