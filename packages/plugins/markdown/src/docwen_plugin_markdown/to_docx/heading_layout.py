"""Retain template numbering indentation when emitting literal heading labels."""

from __future__ import annotations

from typing import Any

from docx.oxml.ns import qn

from docwen_core.docx_parsing.numbering_index import NumberingIndex


def _style_indent(style: Any) -> dict[str, str]:
    """Collect indentation declared by a style chain, nearest style first."""
    collected: dict[str, str] = {}
    seen: set[str] = set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        indent = style.element.find(f"{qn('w:pPr')}/{qn('w:ind')}")
        if indent is not None:
            for key, value in indent.attrib.items():
                collected.setdefault(key, value)
        style = style.base_style
    return collected


def suppress_heading_numbering(document: Any, paragraph: Any) -> None:
    """Disable duplicate labels without dropping inherited indentation.

    A numPr paragraph takes its indentation from the numbering level and stops
    using the style's own indent. Materialize the effective indentation on the
    new paragraph, never on a shared template style: first from the resolved
    numbering level, otherwise from the style chain. Explicit paragraph
    indentation wins.
    """
    style = paragraph.style
    try:
        level = NumberingIndex(document).lookup_by_style_id(style.style_id) if style is not None else None
    except (ValueError, TypeError):
        # A malformed unused template definition must not block conversion.
        level = None
    applied_from_numbering = False
    if level is not None:
        root = document.part.numbering_part.element
        definitions = root.findall(qn("w:abstractNum"))
        abstract = next((n for n in definitions if n.get(qn("w:abstractNumId")) == level.abstract_num_id), None)
        native_level = None
        if abstract is not None:
            native_level = next(
                (n for n in abstract.findall(qn("w:lvl")) if n.get(qn("w:ilvl")) == str(level.ilvl)), None
            )
        instance = next((n for n in root.findall(qn("w:num")) if n.get(qn("w:numId")) == level.num_id), None)
        if instance is not None:
            for override in instance.findall(qn("w:lvlOverride")):
                if override.get(qn("w:ilvl")) == str(level.ilvl):
                    replacement = override.find(qn("w:lvl"))
                    if replacement is not None:
                        native_level = replacement
                    break
        indent = native_level.find(f"{qn('w:pPr')}/{qn('w:ind')}") if native_level is not None else None
        if indent is not None:
            inherited: set[str] = set()
            seen: set[str] = set()
            while style is not None and style.style_id not in seen:
                seen.add(style.style_id)
                style_indent = style.element.find(f"{qn('w:pPr')}/{qn('w:ind')}")
                if style_indent is not None:
                    inherited.update(style_indent.attrib)
                style = style.base_style
            direct = paragraph._p.get_or_add_pPr().get_or_add_ind()
            inherited.update(direct.attrib)
            # The alternative OOXML representations of an indentation dimension
            # must not override an explicit value from a template style.
            groups = (
                {qn(f"w:{name}") for name in ("left", "leftChars", "start", "startChars")},
                {qn(f"w:{name}") for name in ("right", "rightChars", "end", "endChars")},
                {qn(f"w:{name}") for name in ("firstLine", "firstLineChars", "hanging", "hangingChars")},
            )
            for key, value in indent.attrib.items():
                group = next((group for group in groups if key in group), {key})
                if not inherited.intersection(group):
                    direct.set(key, value)
                    applied_from_numbering = True
    if not applied_from_numbering:
        # No numbering definition supplies the indent (the shipped template has no
        # numbering part at all). Copy the style indent onto the paragraph so the
        # numPr below cannot drop it in Word.
        style_indent = _style_indent(style)
        if style_indent:
            direct = paragraph._p.get_or_add_pPr().get_or_add_ind()
            for key, value in style_indent.items():
                if key not in direct.attrib:
                    direct.set(key, value)
    properties = paragraph._p.get_or_add_pPr().get_or_add_numPr()
    properties.get_or_add_ilvl().val = 0
    properties.get_or_add_numId().val = 0
