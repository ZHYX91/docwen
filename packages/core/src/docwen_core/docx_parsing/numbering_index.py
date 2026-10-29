"""Shared Word numbering.xml index for numPr → list-type/level resolution."""

from __future__ import annotations

import re
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
    lvl_text: str | None
    start: int = 1
    suff: str = "tab"
    p_style: str | None = None


class NumberingIndex:
    """Parse ``word/numbering.xml`` to resolve numId → abstractNumId → level."""

    def __init__(self, doc) -> None:
        self._num_to_abstract: dict[str, str] = {}  # numId → abstractNumId
        self._abstract_levels: dict[str, dict[int, dict]] = {}  # abstractNumId → {ilvl: lvl_elem_attrs}
        self._abstract_num_style_links: dict[str, str] = {}  # abstractNumId → numStyleLink pStyle
        self._num_level_overrides: dict[str, dict[int, dict]] = {}
        self._style_numbering: dict[str, tuple[str | None, str | None, str | None]] = {}
        self._style_lookup_cache: dict[str, NumberingLevel | None] = {}

        namespace = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        styles = getattr(getattr(doc, "styles", None), "element", None)
        if styles is not None:
            for style in styles.findall(f"{namespace}style"):
                style_id = style.get(f"{namespace}styleId")
                if not style_id:
                    continue
                properties = style.find(f"{namespace}pPr")
                num_pr = properties.find(f"{namespace}numPr") if properties is not None else None
                num_id = num_pr.find(f"{namespace}numId") if num_pr is not None else None
                ilvl = num_pr.find(f"{namespace}ilvl") if num_pr is not None else None
                based_on = style.find(f"{namespace}basedOn")
                self._style_numbering[style_id] = (
                    num_id.get(f"{namespace}val") if num_id is not None else None,
                    ilvl.get(f"{namespace}val", "0") if ilvl is not None else None,
                    based_on.get(f"{namespace}val") if based_on is not None else None,
                )

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
                    overrides = {}
                    namespace = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
                    for override in num_elem.findall(f"{namespace}lvlOverride"):
                        raw_level = override.get(f"{namespace}ilvl")
                        if raw_level is None:
                            continue
                        values = {}
                        native_level = override.find(f"{namespace}lvl")
                        if native_level is not None:
                            for child_name, key in (
                                ("numFmt", "numFmt"),
                                ("lvlText", "lvlText"),
                                ("start", "start"),
                                ("suff", "suff"),
                                ("pStyle", "pStyle"),
                            ):
                                child = native_level.find(f"{namespace}{child_name}")
                                if child is not None and (value := child.get(f"{namespace}val")) is not None:
                                    values[key] = int(value) if key == "start" else value
                        start_override = override.find(f"{namespace}startOverride")
                        if start_override is not None and (value := start_override.get(f"{namespace}val")) is not None:
                            values["start"] = int(value)
                        overrides[int(raw_level)] = values
                    self._num_level_overrides[str(num_id)] = overrides

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
                lvl_text = None
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
        lvl = {**lvl, **self._num_level_overrides.get(str(num_id), {}).get(ilvl, {})}
        return NumberingLevel(
            num_id=str(num_id),
            abstract_num_id=abs_id,
            ilvl=ilvl,
            num_fmt=lvl.get("numFmt", ""),
            lvl_text=lvl.get("lvlText"),
            start=lvl.get("start", 1),
            suff=lvl.get("suff", "tab"),
            p_style=lvl.get("pStyle") or None,
        )

    def lookup_by_style_id(self, style_id: str) -> NumberingLevel | None:
        """Resolve once per style in this request-owned numbering snapshot."""
        if style_id not in self._style_lookup_cache:
            self._style_lookup_cache[style_id] = self._lookup_style_binding(style_id)
        return self._style_lookup_cache[style_id]

    def _lookup_style_binding(self, style_id: str) -> NumberingLevel | None:
        """Resolve a concrete style binding, never guess the first instance."""
        ancestors: set[str] = set()
        style_chain: list[str] = []
        current: str | None = style_id
        bound_num_id: str | None = None
        bound_level: int | None = None
        while current is not None and current not in ancestors:
            ancestors.add(current)
            style_chain.append(current)
            num_id, raw_ilvl, current = self._style_numbering.get(current, (None, None, None))
            if bound_num_id is None and num_id is not None:
                bound_num_id = num_id
            if bound_level is None and raw_ilvl is not None:
                try:
                    bound_level = int(raw_ilvl)
                except ValueError:
                    # An invalid used style has no usable concrete binding;
                    # an unrelated unused style never blocks this document.
                    return None
        if bound_num_id == "0":
            return None
        if bound_num_id is not None:
            if bound_level is not None:
                return self.lookup(bound_num_id, bound_level)
            levels = self._abstract_levels.get(self._num_to_abstract.get(bound_num_id, ""), {})
            for candidate_style in style_chain:
                candidates = [
                    effective
                    for ilvl in levels
                    if (effective := self.lookup(bound_num_id, ilvl)) is not None
                    and effective.p_style == candidate_style
                ]
                if candidates:
                    return candidates[0] if len(candidates) == 1 else None
            return self.lookup(bound_num_id, 0)

        for candidate_style in style_chain:
            candidates = [
                effective
                for num_id, abs_id in self._num_to_abstract.items()
                for ilvl in self._abstract_levels.get(abs_id, {})
                if (effective := self.lookup(num_id, ilvl)) is not None and effective.p_style == candidate_style
            ]
            if candidates:
                return candidates[0] if len(candidates) == 1 else None
        return None

    def lookup_by_num_id(self, num_id: str, ilvl: int) -> NumberingLevel | None:
        """Direct lookup by numId and ilvl (convenience alias for ``lookup``)."""
        return self.lookup(num_id, ilvl)

    def preview_numbering_text(self, level: NumberingLevel, ordinal: int = 1) -> str:
        return _format_counter(ordinal, level.num_fmt)

    def render_numbering_text(
        self,
        level: NumberingLevel,
        counter_value: int,
        parent_counters: dict[int, int] | None = None,
    ) -> str:
        """Resolve ``lvlText`` placeholders against counter values.

        Substitutes applicable ``%1``..``%9`` placeholders using the same
        concrete instance, with *counter_value* for the current level.

        Args:
            level: The ``NumberingLevel`` to render.
            counter_value: The counter value for *level.ilvl*.
            parent_counters: Optional ``{ilvl: value}`` map for sibling levels.

        Returns:
            The resolved numbering prefix text (e.g. ``"1.1 "``, ``"一、"``).
        """
        lvl_text = level.lvl_text
        if lvl_text is None:
            return _apply_numbering_suffix(self.preview_numbering_text(level, counter_value), level.suff)

        parents = parent_counters or {}

        def replace_placeholder(match: re.Match[str]) -> str:
            ilvl = int(match.group(1)) - 1
            if ilvl > level.ilvl:
                # OOXML ignores a level reference higher than this level;
                # literal punctuation in the template remains authored text.
                return ""
            effective = (
                level
                if ilvl == level.ilvl
                else (
                    self.lookup(level.num_id, ilvl)
                    if level.num_id
                    else self.lookup_by_abstract(level.abstract_num_id, ilvl)
                )
            )
            value = counter_value if ilvl == level.ilvl else parents.get(ilvl)
            if value is None:
                value = effective.start if effective is not None else 1
            return _format_counter(value, effective.num_fmt if effective is not None else "decimal")

        return _apply_numbering_suffix(re.sub(r"%([1-9])", replace_placeholder, lvl_text), level.suff)

    def lookup_by_abstract(self, abstract_num_id: str, ilvl: int) -> NumberingLevel | None:
        """Look up a numbering level by abstractNumId and ilvl."""
        levels = self._abstract_levels.get(abstract_num_id, {})
        lvl = levels.get(ilvl)
        if lvl is None:
            return None
        return NumberingLevel(
            num_id="",
            abstract_num_id=abstract_num_id,
            ilvl=ilvl,
            num_fmt=lvl.get("numFmt", ""),
            lvl_text=lvl.get("lvlText"),
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
