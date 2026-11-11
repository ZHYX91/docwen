"""Section ownership while rendered blocks are staged outside their template slot."""

from __future__ import annotations

from copy import deepcopy
from typing import Any

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

_SECTION_TYPES = {"next": "nextPage", "continuous": "continuous", "even": "evenPage", "odd": "oddPage"}


def section_properties_at(document: Any, paragraph: Any | None = None) -> Any | None:
    """Find the owning boundary, or None for an implicit final section."""
    body = document.element.body
    if paragraph is None:
        return body.sectPr
    element = paragraph._element if hasattr(paragraph, "_element") else paragraph
    if element.getparent() is not body or element.tag != qn("w:p"):
        raise ValueError("A section boundary must belong to a main-document paragraph")
    for sibling in list(body)[body.index(element) :]:
        if sibling.tag == qn("w:sectPr"):
            return sibling
        if sibling.tag == qn("w:p"):
            boundary = sibling.find(f"{qn('w:pPr')}/{qn('w:sectPr')}")
            if boundary is not None:
                return boundary
    return None


class SectionBreakContext:
    """Keep generated breaks attached to the section owning the insertion slot."""

    def __init__(self, document: Any, anchor: Any | None = None) -> None:
        self.document = document
        self.boundary = section_properties_at(document, anchor)

    def append_break(self, paragraph: Any, section_type: str = "next") -> None:
        element = paragraph._element
        if element.getparent() is not self.document.element.body:
            raise ValueError("A section break cannot be inserted outside the main document body")
        properties = element.get_or_add_pPr()
        if properties.find(qn("w:sectPr")) is not None:
            raise ValueError("A paragraph already owns a section boundary")
        if self.boundary is None:
            # body/sectPr is optional. Preserve that omission until a requested
            # split actually requires explicit section boundaries.
            self.boundary = self.document.element.body.get_or_add_sectPr()
        # sectPr describes the content preceding it, but w:type describes how
        # that section starts relative to its predecessor (ISO 29500 17.6.22).
        # The former section retains its settings; the following section gets
        # the requested start type. Copy complete XML, including extensions.
        properties.insert_element_before(deepcopy(self.boundary), "w:pPrChange")
        start = self.boundary.get_or_add_type()
        start.set(qn("w:val"), _SECTION_TYPES.get(section_type, "nextPage"))
        numbering = self.boundary.find(qn("w:pgNumType"))
        if numbering is not None:
            # The authored restart belongs to the original section only.
            # Its number format and chapter-number settings still apply.
            numbering.attrib.pop(qn("w:start"), None)


_RANGE_MARKERS = frozenset(
    qn(f"w:{name}{suffix}")
    for name in (
        "bookmark",
        "commentRange",
        "customXmlDelRange",
        "customXmlInsRange",
        "customXmlMoveFromRange",
        "customXmlMoveToRange",
        "moveFromRange",
        "moveToRange",
        "perm",
    )
    for suffix in ("Start", "End")
)


def splice_body_elements(document: Any, elements: list[Any], anchor: Any | None) -> None:
    """Replace a body slot without losing its section or attached range identity."""
    body = document.element.body
    elements = list(elements)
    placeholder = anchor._element if anchor is not None else None
    if placeholder is not None and (placeholder.getparent() is not body or placeholder.tag != qn("w:p")):
        raise ValueError("The body placeholder is no longer a main-document paragraph")
    if placeholder in elements or len({id(element) for element in elements}) != len(elements):
        raise ValueError("Rendered body elements must be distinct from their template anchor")
    for element in elements:
        parent = element.getparent()
        if parent is not None:
            parent.remove(element)
    if placeholder is None:
        terminal = body.sectPr
        index = body.index(terminal) if terminal is not None else len(body)
        for offset, element in enumerate(elements):
            body.insert(index + offset, element)
        return

    atoms = [atom for child in placeholder if child.tag != qn("w:pPr") for atom in _inline_atoms(child)]
    last_text = max(
        (index for index, (node, _wrappers, is_text) in enumerate(atoms) if is_text and (node.text or "").strip()),
        default=-1,
    )
    # Consume only visible marker text. Unknown subtrees are opaque. Retained
    # nodes inside a split marker share its leading edge, keeping their entire
    # original order, including mixed references, ranges and unknown content.
    prefix = _wrap_atoms([(node, wrappers) for node, wrappers, is_text in atoms[: last_text + 1] if not is_text])
    suffix = _wrap_atoms([(node, wrappers) for node, wrappers, is_text in atoms[last_text + 1 :] if not is_text])

    properties = placeholder.find(qn("w:pPr"))
    boundary = properties.find(qn("w:sectPr")) if properties is not None else None
    _retain_edge(elements, prefix, properties, before=True)
    _retain_edge(elements, suffix, properties, before=False)
    if boundary is not None:
        last = next((element for element in reversed(elements) if element.tag not in _RANGE_MARKERS), None)
        if last is None or last.tag != qn("w:p") or last.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None:
            last = _carrier_paragraph(properties)
            elements.append(last)
        last.get_or_add_pPr().insert_element_before(boundary, "w:pPrChange")

    index = body.index(placeholder)
    body.remove(placeholder)
    for offset, element in enumerate(elements):
        body.insert(index + offset, element)


def _inline_atoms(child: Any, wrappers: tuple[Any, ...] = ()) -> list[tuple[Any, tuple[Any, ...], bool]]:
    if child.tag in _RANGE_MARKERS:
        # A range inside a hyperlink must enclose the replacement blocks too;
        # range markers permit paragraph/body parents and need no link wrapper.
        return [(child, (), False)]
    if child.tag == qn("w:r"):
        return [(node, (*wrappers, child), node.tag == qn("w:t")) for node in child if node.tag != qn("w:rPr")]
    if child.tag == qn("w:hyperlink"):
        atoms = [atom for node in child for atom in _inline_atoms(node, (*wrappers, child))]
        if not any(not is_text and any(wrapper is child for wrapper in parents) for _, parents, is_text in atoms):
            # Preserve link attributes/relationship identity after its marker
            # text is consumed, without retaining the old text or adding a line.
            atoms.append((_inline_shell(child), wrappers, False))
        return atoms
    return [(child, wrappers, False)]


def _inline_shell(element: Any) -> Any:
    shell = deepcopy(element)
    for child in list(shell):
        if element.tag != qn("w:r") or child.tag != qn("w:rPr"):
            shell.remove(child)
    return shell


def _wrap_atoms(atoms: list[tuple[Any, tuple[Any, ...]]]) -> list[Any]:
    result: list[Any] = []
    previous: tuple[Any, ...] = ()
    containers: list[Any] = []
    for node, wrappers in atoms:
        shared = 0
        while shared < min(len(previous), len(wrappers)) and previous[shared] is wrappers[shared]:
            shared += 1
        containers = containers[:shared]
        for original in wrappers[shared:]:
            shell = _inline_shell(original)
            if containers:
                containers[-1].append(shell)
            else:
                result.append(shell)
            containers.append(shell)
        if containers:
            containers[-1].append(node)
        else:
            result.append(node)
        previous = wrappers
    return result


def _carrier_paragraph(properties: Any | None) -> Any:
    paragraph = OxmlElement("w:p")
    if properties is not None:
        copied = deepcopy(properties)
        for boundary in copied.findall(qn("w:sectPr")):
            copied.remove(boundary)
        paragraph.append(copied)
    return paragraph


def _retain_edge(elements: list[Any], nodes: list[Any], properties: Any | None, *, before: bool) -> None:
    if not nodes:
        return
    candidate = (elements[0] if before else elements[-1]) if elements else None
    metadata_only = all(_range_metadata(node) for node in nodes)
    if (
        metadata_only
        and candidate is not None
        and candidate.tag == qn("w:p")
        and (before or candidate.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is None)
    ):
        index = int(candidate.find(qn("w:pPr")) is not None) if before else len(candidate)
        for offset, node in enumerate(nodes):
            candidate.insert(index + offset, node)
        return
    if all(node.tag in _RANGE_MARKERS for node in nodes):
        # These range markers also allow body as a parent (ISO 29500 17.13).
        edge = nodes
    else:
        carrier = _carrier_paragraph(properties)
        carrier.extend(nodes)
        edge = [carrier]
    if before:
        elements[:0] = edge
    else:
        elements.extend(edge)


def _range_metadata(node: Any) -> bool:
    if node.tag in _RANGE_MARKERS:
        return True
    if node.tag == qn("w:r"):
        return all(child.tag in {qn("w:rPr"), qn("w:commentReference")} for child in node)
    if node.tag == qn("w:hyperlink"):
        return all(_range_metadata(child) for child in node)
    return False


__all__ = ["SectionBreakContext", "section_properties_at", "splice_body_elements"]
