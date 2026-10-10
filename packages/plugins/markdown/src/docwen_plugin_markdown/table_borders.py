"""Project multi-row three-line headers onto native, already merged cells."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from typing import Any

from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def table_header_separator(style: Any, *, fallback: Mapping[str, str] | None = None) -> dict[str, str] | None:
    """Read the selected style's header rule without inventing template defaults.

    An explicit none/nil opts out of this projection and is preserved as
    authored. Word does not treat those values alike, so neither is converted
    to the other or advertised as equivalent native border suppression.
    A missing rule may inherit through basedOn; an unresolved/cyclic chain or
    ambiguous conditional rule grants no permission to add a separator.
    Other authored horizontal cell edges, insideH or nonzero cell spacing
    opt out: nil must not suppress an independent template border intent.
    Only a caller-owned direct fallback may supply a missing separator; it
    still respects every authored rule and conflict in the default style.
    """
    seen: set[str] = set()
    separator = None
    while style is not None:
        if style.style_id in seen:
            return None
        seen.add(style.style_id)
        scopes = [style.element, *style.element.findall(qn("w:tblStylePr"))]
        for scope in scopes:
            if scope.find(f"{qn('w:tblPr')}/{qn('w:tblBorders')}/{qn('w:insideH')}") is not None:
                return None
            spacing = scope.find(f"{qn('w:tblPr')}/{qn('w:tblCellSpacing')}")
            if spacing is not None and spacing.get(qn("w:w")) != "0":
                return None
            borders = scope.find(f"{qn('w:tcPr')}/{qn('w:tcBorders')}")
            if borders is not None:
                allowed_bottom = scope.get(qn("w:type")) == "firstRow"
                if any(
                    borders.find(qn(f"w:{edge}")) is not None
                    for edge in (("top", "insideH") if allowed_bottom else ("top", "bottom", "insideH"))
                ):
                    return None
        rules = [rule for rule in style.element.findall(qn("w:tblStylePr")) if rule.get(qn("w:type")) == "firstRow"]
        if len(rules) > 1:
            return None
        if rules and separator is None:
            bottom = rules[0].find(f"{qn('w:tcPr')}/{qn('w:tcBorders')}/{qn('w:bottom')}")
            if bottom is not None:
                value = bottom.get(qn("w:val"), "")
                if not value or value in {"none", "nil"}:
                    return None
                separator = dict(bottom.attrib)
        style = style.base_style
    return separator if separator is not None else (dict(fallback) if fallback is not None else None)


def apply_multirow_header_borders(
    table: Any,
    *,
    header_rows: int,
    anchors: Sequence[Mapping[str, Any]],
    separator: Mapping[str, str] | None,
) -> None:
    """Write only rectangle bottom edges, after gridSpan/vMerge are final.

    The complete header boundary and wide groups ending above it get the
    selected style's separator. Other header rectangle bottoms get nil to
    suppress the old firstRow condition, except for an authored direct edge.
    Interior vMerge edges are left alone: a nil there could suppress the real
    merged rectangle's separator. No table/style/repetition property changes.
    """
    if header_rows <= 1 or separator is None:
        return
    edges: dict[tuple[int, int, int], bool] = {}
    for anchor in anchors:
        row = int(anchor["row"])
        end = row + int(anchor["row_span"])
        if row >= header_rows or end > header_rows:
            continue
        column = int(anchor["column"])
        span = int(anchor["column_span"])
        edges[(end - 1, column, span)] = end == header_rows or span > 1

    for row_index, row in enumerate(table.findall(qn("w:tr"))[:header_rows]):
        column = 0
        for cell in row.findall(qn("w:tc")):
            properties = cell.find(qn("w:tcPr"))
            span_node = properties.find(qn("w:gridSpan")) if properties is not None else None
            span = int(span_node.get(qn("w:val"), "1")) if span_node is not None else 1
            key = (row_index, column, span)
            column += span
            if key not in edges:
                continue
            if properties is None:
                properties = OxmlElement("w:tcPr")
                cell.insert(0, properties)
            borders = properties.find(qn("w:tcBorders"))
            if borders is None:
                borders = OxmlElement("w:tcBorders")
                properties.append(borders)
            if borders.find(qn("w:bottom")) is not None:
                continue
            bottom = OxmlElement("w:bottom")
            for name, value in (separator if edges[key] else {qn("w:val"): "nil"}).items():
                bottom.set(name, value)
            borders.append(bottom)
