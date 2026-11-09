"""Retain template numbering indentation when emitting literal heading labels."""

from __future__ import annotations

from typing import Any

from docx.oxml.ns import qn

from docwen_core.docx_parsing.numbering_index import NumberingIndex

# Legacy left/right and newer start/end describe the same logical edges:
# left/start is the leading edge even in RTL (ECMA-376 Part 4, 2.3.1.12).
# Keep these aliases together; do not swap them based on paragraph bidi.
_INDENT_DIMENSIONS = (
    frozenset(qn(f"w:{name}") for name in ("left", "leftChars", "start", "startChars")),
    frozenset(qn(f"w:{name}") for name in ("right", "rightChars", "end", "endChars")),
    frozenset(qn(f"w:{name}") for name in ("firstLine", "firstLineChars", "hanging", "hangingChars")),
)


def _inherit_indent(target: dict[str, str], source: Any) -> None:
    """Fill missing dimensions without mixing mutually exclusive representations.

    Keep all attributes from the winning layer together, so Word can apply its
    character/twip and first-line/hanging precedence within that layer.
    """
    occupied = set(target)
    for key, value in source.items():
        dimension = next((group for group in _INDENT_DIMENSIONS if key in group), {key})
        if not occupied.intersection(dimension):
            target[key] = value


def _style_indent(style: Any) -> dict[str, str]:
    """Collect indentation dimensions from the nearest style declaring each one."""
    collected: dict[str, str] = {}
    seen: set[str] = set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        indent = style.element.find(f"{qn('w:pPr')}/{qn('w:ind')}")
        if indent is not None:
            _inherit_indent(collected, indent.attrib)
        style = style.base_style
    return collected


def suppress_heading_numbering(document: Any, paragraph: Any) -> None:
    """Disable duplicate labels without dropping inherited indentation.

    Materialize each indentation dimension on the new paragraph, never on a
    shared template style. Explicit paragraph values win, followed by the nearest
    style declaration, then the resolved numbering level for missing dimensions.
    """
    style = paragraph.style
    try:
        level = NumberingIndex(document).lookup_by_style_id(style.style_id) if style is not None else None
    except (ValueError, TypeError):
        # A malformed unused template definition must not block conversion.
        level = None
    properties = paragraph._p.get_or_add_pPr()
    direct = properties.find(qn("w:ind"))
    effective = dict(direct.attrib) if direct is not None else {}
    _inherit_indent(effective, _style_indent(style))
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
            _inherit_indent(effective, indent.attrib)
    if effective:
        direct = properties.get_or_add_ind()
        for key, value in effective.items():
            direct.set(key, value)
    properties = paragraph._p.get_or_add_pPr().get_or_add_numPr()
    properties.get_or_add_ilvl().val = 0
    properties.get_or_add_numId().val = 0
