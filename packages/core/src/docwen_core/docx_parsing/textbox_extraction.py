"""Extract textbox paragraphs from DOCX body, headers, and footers.

Handles DrawingML, VML, and ``mc:AlternateContent`` Choice/Fallback
structures by XML occurrence, selecting one alternate representation.
"""

from __future__ import annotations

from dataclasses import dataclass

from docwen_core.docx_parsing.xml_ns import (
    NS_MC,
    NS_V,
    NS_W,
    NS_WPS,
    NSMAP_WPV,
)


@dataclass(frozen=True)
class ExtractedParagraph:
    text: str
    anchor_index: int | None  # element order anchor in body
    source: str  # textbox, header_textbox, footer_textbox
    part_name: str  # document, headerN, footerN


def extract_textbox_paragraphs(doc) -> list[ExtractedParagraph]:
    """Extract textbox paragraphs from document body, headers, and footers.

    Returns a flat list of ``ExtractedParagraph`` objects in document order.
    """
    results: list[ExtractedParagraph] = []

    # Body
    results.extend(
        _extract_from_part(doc.element.body, doc.part, anchor_offset=0, source="textbox", part_name="document")
    )

    # Headers
    hdr_parts = getattr(doc.part, "header_parts", []) or []
    for i, hdr_part in enumerate(hdr_parts):
        try:
            body_elem = hdr_part.element
        except AttributeError:
            continue
        results.extend(
            _extract_from_part(body_elem, hdr_part, anchor_offset=None, source="header_textbox", part_name=f"header{i}")
        )

    # Footers
    ftr_parts = getattr(doc.part, "footer_parts", []) or []
    for i, ftr_part in enumerate(ftr_parts):
        try:
            body_elem = ftr_part.element
        except AttributeError:
            continue
        results.extend(
            _extract_from_part(body_elem, ftr_part, anchor_offset=None, source="footer_textbox", part_name=f"footer{i}")
        )

    return results


def _extract_from_part(
    body_elem,
    part,
    *,
    anchor_offset: int | None,
    source: str,
    part_name: str,
) -> list[ExtractedParagraph]:
    """Extract textbox paragraphs from one part (body / header / footer)."""
    results: list[ExtractedParagraph] = []
    # Keep the element wrappers themselves alive while deduplicating. Storing
    # only ``id(element)`` lets lxml wrapper objects be reclaimed between
    # textboxes, after which CPython may reuse the id for a different paragraph
    # and silently discard valid content.
    processed_elements: set[object] = set()

    def walk(element) -> list[ExtractedParagraph]:
        if element.tag == f"{{{NS_MC}}}AlternateContent":
            # Choice/Fallback are representations of one shape, not separate
            # authored occurrences. Select the first supported nonempty Choice;
            # only use Fallback when no Choice provides textbox paragraphs.
            for choice in element.findall(f"{{{NS_MC}}}Choice", NSMAP_WPV):
                selected = walk(choice)
                if selected:
                    return selected
            fallback = element.find(f"{{{NS_MC}}}Fallback", NSMAP_WPV)
            return walk(fallback) if fallback is not None else []
        if element.tag in {f"{{{NS_WPS}}}txbx", f"{{{NS_V}}}textbox"}:
            return _extract_from_txbx(
                element,
                part,
                _resolve_anchor_index(body_elem, element, anchor_offset),
                source,
                part_name,
                processed_elements,
            )
        return [paragraph for child in element for paragraph in walk(child)]

    results.extend(walk(body_elem))
    return results


def _resolve_anchor_index(body_elem, descendant, anchor_offset: int | None) -> int | None:
    """Return the direct-child index that owns *descendant* in a body part."""
    if anchor_offset is None:
        return None

    current = descendant
    parent = current.getparent() if hasattr(current, "getparent") else None
    while parent is not None and parent is not body_elem:
        current = parent
        parent = current.getparent() if hasattr(current, "getparent") else None

    if parent is not body_elem:
        return None
    return anchor_offset + body_elem.index(current)


def _extract_from_txbx(
    txbx,
    part,
    anchor_offset,
    source,
    part_name,
    processed_elements,
) -> list[ExtractedParagraph]:
    """Extract each real paragraph occurrence from a DrawingML or VML box."""
    results: list[ExtractedParagraph] = []
    for para_elem in txbx.findall(f".//{{{NS_W}}}p", NSMAP_WPV):
        if para_elem in processed_elements:
            continue
        processed_elements.add(para_elem)
        text = _extract_para_text(para_elem, part)
        if text:
            results.append(
                ExtractedParagraph(
                    text=text,
                    anchor_index=anchor_offset,
                    source=source,
                    part_name=part_name,
                )
            )
    return results


def _extract_para_text(para_elem, part) -> str:
    """Project one textbox paragraph using Word's accepted visible view."""

    def _on_off_property_is_enabled(element) -> bool:
        value = element.get(f"{{{NS_W}}}val")
        return value is None or value.lower() not in {"0", "false", "off", "no"}

    def _visible_text(element) -> str:
        local_name = element.tag.split("}")[-1] if "}" in (element.tag or "") else (element.tag or "")
        if local_name in {"del", "moveFrom"}:
            return ""
        if local_name == "r":
            run_properties = element.find(f"{{{NS_W}}}rPr")
            vanish = run_properties.find(f"{{{NS_W}}}vanish") if run_properties is not None else None
            if vanish is not None and _on_off_property_is_enabled(vanish):
                return ""
        if local_name == "t":
            return element.text or ""
        if local_name == "tab":
            return "\t"
        if local_name in {"br", "cr"}:
            return "\n"
        if local_name == "noBreakHyphen":
            return "\u2011"
        if local_name == "softHyphen":
            return "\u00ad"
        return "".join(_visible_text(child) for child in element)

    return _visible_text(para_elem).strip()
