"""Conservative readable projection of ordinary heading hyperlinks.

Tooltip and bookmark text are untrusted document content, not authenticated
recovery metadata. Frozen semantic bookmarks and block-anchor SDTs are never
used as navigation authority here.
"""

from __future__ import annotations

import re
from typing import Any
from urllib.parse import quote, unquote

from docx.oxml.ns import qn

from docwen_core.docx_bookmarks import build_docx_bookmark_inventory, prove_bookmark_name

_NAVIGATION_NAME = re.compile(r"DWNav_[0-9a-f]{32}")
_IGNORABLE = frozenset(qn(tag) for tag in ("w:pPr", "w:bookmarkStart", "w:bookmarkEnd", "w:proofErr"))


def _heading_level(element: Any, styles: dict[str, list[Any]]) -> int | None:
    """Require a real outline level, not a heading-like style name."""
    properties = element.find(qn("w:pPr"))
    visited: set[str] = set()
    while True:
        if properties is not None:
            outlines = properties.findall(qn("w:outlineLvl"))
            if outlines:
                if len(outlines) != 1:
                    return None
                value = outlines[0].get(qn("w:val"), "")
                return int(value) if value in tuple(str(level) for level in range(9)) else None
        if element.tag == qn("w:p"):
            references = properties.findall(qn("w:pStyle")) if properties is not None else []
        else:
            references = element.findall(qn("w:basedOn"))
        if len(references) != 1:
            return None
        style_id = references[0].get(qn("w:val"), "")
        matches = styles.get(style_id, [])
        if style_id in visited or len(matches) != 1:
            return None
        visited.add(style_id)
        element = matches[0]
        if element.get(qn("w:type")) != "paragraph":
            return None
        properties = element.find(qn("w:pPr"))


def _heading_text(element: Any) -> str | None:
    """Read only unambiguous visible runs; unknown constructs do not prove a title."""
    if any(text.strip() for text in element.xpath("text()")):
        return None
    parts: list[str] = []
    for child in element:
        if child.tag in _IGNORABLE or child.tag == qn("w:rPr"):
            continue
        if child.tag == qn("w:r"):
            properties = child.find(qn("w:rPr"))
            if properties is not None:
                hidden = properties.find(qn("w:vanish"))
                if hidden is not None and hidden.get(qn("w:val"), "true").lower() not in {"0", "false", "off"}:
                    continue
            text = _heading_text(child)
        elif child.tag in {qn("w:hyperlink"), qn("w:ins"), qn("w:moveTo")}:
            text = _heading_text(child)
        elif child.tag in {qn("w:del"), qn("w:moveFrom")}:
            continue
        elif child.tag == qn("w:t"):
            if len(child):
                return None
            text = child.text or ""
        elif child.tag in {qn("w:br"), qn("w:cr")}:
            if child.get(qn("w:type"), "textWrapping") != "textWrapping":
                return None
            text = " "
        elif child.tag == qn("w:tab"):
            text = "\t"
        else:
            return None
        if text is None:
            return None
        parts.append(text)
    return "".join(parts)


def _properties_hide_text(properties: Any) -> bool:
    if properties is None:
        return False
    return any(
        item.get(qn("w:val"), "true").lower() not in {"0", "false", "off"}
        for tag in ("w:vanish", "w:webHidden", "w:specVanish")
        for item in properties.findall(qn(tag))
    )


def _style_may_hide_text(style_id: str, styles: dict[str, list[Any]]) -> bool:
    visited: set[str] = set()
    while style_id:
        matches = styles.get(style_id, [])
        if style_id in visited or len(matches) != 1:
            return True
        visited.add(style_id)
        style = matches[0]
        if _properties_hide_text(style.find(qn("w:rPr"))):
            return True
        parents = style.findall(qn("w:basedOn"))
        if not parents:
            return False
        if len(parents) != 1:
            return True
        style_id = parents[0].get(qn("w:val"), "")
    return True


def _heading_may_hide_text(element: Any, styles: dict[str, list[Any]]) -> bool:
    # Hidden style inheritance is not private recovery evidence. Reject any
    # uncertain visibility rather than matching text that Word may not show.
    references = list(element.iter(qn("w:pStyle"))) + list(element.iter(qn("w:rStyle")))
    if any(_style_may_hide_text(item.get(qn("w:val"), ""), styles) for item in references):
        return True
    if not element.findall(f"{qn('w:pPr')}/{qn('w:pStyle')}"):
        for candidates in styles.values():
            for style in candidates:
                if (
                    style.get(qn("w:type")) == "paragraph"
                    and style.get(qn("w:default")) in {"1", "true", "on"}
                    and _style_may_hide_text(style.get(qn("w:styleId"), ""), styles)
                ):
                    return True
    return any(_properties_hide_text(properties) for properties in element.iter(qn("w:rPr")))


def _matches_heading(title: str, fragment: str) -> bool:
    title = title.strip()
    slug = re.sub(r"[^\w\- ]", "", title.lower()).replace(" ", "-")
    return bool(title) and fragment in {title, slug}


def resolve_ordinary_heading_target(paragraph: Any, hyperlink: Any) -> str | None:
    """Return a safe fragment only after structural and visible-heading checks."""
    try:
        anchor = hyperlink.get(qn("w:anchor"), "")
        hint = hyperlink.get(qn("w:tooltip"), "")
        if _NAVIGATION_NAME.fullmatch(anchor) is None or not hint.startswith("#"):
            return None
        fragment = unquote(hint[1:], errors="strict")
        if (
            not fragment
            or fragment.startswith("^")
            or any(ord(char) < 32 or ord(char) == 127 for char in hint + fragment)
        ):
            return None
        document = paragraph.part.package.main_document_part.document
        body = document.element.body
        inventory = build_docx_bookmark_inventory(document)
        proof = prove_bookmark_name(inventory, anchor, scope_element=body)
        if not proof.valid or proof.start is None or proof.end is None:
            return None
        start, end = proof.start.element, proof.end.element
        target = start.getparent()
        if (
            target is None
            or target.tag != qn("w:p")
            or end.getparent() is not target
            or proof.start.name != anchor
            or proof.start.raw_id != proof.end.raw_id
            or set(start.attrib) != {qn("w:id"), qn("w:name")}
            or set(end.attrib) != {qn("w:id")}
            or len(start)
            or len(end)
            or start.text
            or end.text
        ):
            return None
        # Frozen bookmarks may coexist but never prove this ordinary target.
        # A second ordinary navigation range makes its ownership ambiguous.
        if any(
            marker is not start and (marker.get(qn("w:name")) or "").startswith("DWNav_")
            for marker in target.iter(qn("w:bookmarkStart"))
        ):
            return None
        # A moved or partial range cannot stand in for a whole heading.
        children = list(target)
        first, last = children.index(start), children.index(end)
        if any(child.tag not in _IGNORABLE for child in children[:first] + children[last + 1 :]):
            return None
        # Only the main story may supply targets, never textboxes or deleted text.
        owner = target.getparent()
        while owner is not body:
            if owner is None or owner.tag not in {qn("w:tbl"), qn("w:tr"), qn("w:tc"), qn("w:sdt"), qn("w:sdtContent")}:
                return None
            owner = owner.getparent()
        styles: dict[str, list[Any]] = {}
        for style in document.styles.element:
            if style.tag == qn("w:style"):
                styles.setdefault(style.get(qn("w:styleId"), ""), []).append(style)
        defaults = document.styles.element.find(f"{qn('w:docDefaults')}/{qn('w:rPrDefault')}/{qn('w:rPr')}")
        if _properties_hide_text(defaults):
            return None
        if _heading_level(target, styles) is None:
            return None
        matched: list[Any] = []
        for candidate in body.iter(qn("w:p")):
            if _heading_level(candidate, styles) is None:
                continue
            if _heading_may_hide_text(candidate, styles):
                return None
            title = _heading_text(candidate)
            if title is None:
                return None
            if _matches_heading(title, fragment):
                matched.append(candidate)
        if len(matched) != 1 or matched[0] is not target:
            return None
        return "#" + quote(hint[1:], safe="/%-._~")
    except (AttributeError, KeyError, TypeError, ValueError):
        return None
