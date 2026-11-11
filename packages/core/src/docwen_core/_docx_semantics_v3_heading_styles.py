"""Exact, request-owned Heading identities independent of template style names."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

from docx.opc.constants import RELATIONSHIP_TYPE
from docx.oxml.ns import qn
from lxml import etree

from docwen_core._docx_semantics_v3_model import DocxSemanticsV3Error, require_source_id

HEADING_STYLE_BINDING_MAP_NAMESPACE = "https://docwen.dev/schema/document-heading-style-binding-map/v1"
_XML_STRING = re.compile(r"[\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD\U00010000-\U0010FFFF]+")


@dataclass(frozen=True, slots=True)
class HeadingStyleBindingV3:
    source_id: str
    level: int
    resolved_style_id: str
    visible_name: str


def _validate(bindings: tuple[HeadingStyleBindingV3, ...]) -> tuple[HeadingStyleBindingV3, ...]:
    seen: set[str] = set()
    for binding in bindings:
        require_source_id(binding.source_id)
        if binding.source_id in seen:
            raise DocxSemanticsV3Error("Heading-style bindings have duplicate target IDs")
        seen.add(binding.source_id)
        if type(binding.level) is not int or not 1 <= binding.level <= 9:
            raise DocxSemanticsV3Error("Heading-style binding level must be 1 through 9")
        if any(_XML_STRING.fullmatch(value) is None for value in (binding.resolved_style_id, binding.visible_name)):
            raise DocxSemanticsV3Error("Heading-style identity must contain nonempty XML strings")
    return tuple(sorted(bindings, key=lambda item: item.source_id))


def _style(styles: Any, style_id: str) -> Any:
    matches = [item for item in styles.findall(qn("w:style")) if item.get(qn("w:styleId")) == style_id]
    if len(matches) != 1 or matches[0].get(qn("w:type")) != "paragraph":
        raise DocxSemanticsV3Error("Heading-style binding must resolve one exact paragraph style")
    return matches[0]


def _style_registry(document: Any) -> Any:
    try:
        part = document.part.part_related_by(RELATIONSHIP_TYPE.STYLES)
    except (KeyError, ValueError) as exc:
        raise DocxSemanticsV3Error("Heading-style binding requires one existing styles registry") from exc
    if not hasattr(part, "element") or part.element.tag != qn("w:styles"):
        raise DocxSemanticsV3Error("Heading-style registry has an invalid root")
    return part.element


def capture_heading_style_binding(
    document: Any, paragraph: Any, source_id: str, level: int | None
) -> HeadingStyleBindingV3:
    styles = _style_registry(document)
    references = paragraph._p.findall(f"{qn('w:pPr')}/{qn('w:pStyle')}")
    if len(references) != 1:
        raise DocxSemanticsV3Error("Heading target must have exactly one direct paragraph style")
    style_id = references[0].get(qn("w:val"), "")
    style = _style(styles, style_id)
    names = style.findall(qn("w:name"))
    if len(names) != 1:
        raise DocxSemanticsV3Error("Heading style must have one primary name")
    if level is None:
        # Direct Core callers can provide an actual inherited outline level.
        # Names, aliases and conventional IDs never establish Heading identity.
        current = style
        seen: set[str] = set()
        while current is not None:
            identity = current.get(qn("w:styleId"), "")
            if identity in seen:
                raise DocxSemanticsV3Error("Heading style inheritance contains a cycle")
            seen.add(identity)
            outlines = current.findall(f"{qn('w:pPr')}/{qn('w:outlineLvl')}")
            if len(outlines) > 1:
                raise DocxSemanticsV3Error("Heading style outline level is duplicated")
            if outlines:
                value = outlines[0].get(qn("w:val"), "")
                if value not in {str(index) for index in range(9)}:
                    raise DocxSemanticsV3Error("Heading style has no Heading outline level")
                level = int(value) + 1
                break
            parents = current.findall(qn("w:basedOn"))
            if len(parents) > 1:
                raise DocxSemanticsV3Error("Heading style inheritance is ambiguous")
            current = _style(styles, parents[0].get(qn("w:val"), "")) if parents else None
        if level is None:
            raise DocxSemanticsV3Error("Heading binding requires an explicit level or an inherited outline level")
    binding = HeadingStyleBindingV3(source_id, level, style_id, names[0].get(qn("w:val"), ""))
    return _validate((binding,))[0]


def heading_style_binding_map_xml(bindings: tuple[HeadingStyleBindingV3, ...]) -> bytes:
    namespace = HEADING_STYLE_BINDING_MAP_NAMESPACE
    root = etree.Element(f"{{{namespace}}}documentHeadingStyleBindingMap", nsmap={"dw": namespace}, version="1")
    for binding in _validate(bindings):
        etree.SubElement(
            root,
            f"{{{namespace}}}binding",
            source_id=binding.source_id,
            level=str(binding.level),
            resolved_style_id=binding.resolved_style_id,
            visible_name=binding.visible_name,
        )
    return etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)


def parse_heading_style_binding_map(root: Any) -> tuple[HeadingStyleBindingV3, ...]:
    namespace = f"{{{HEADING_STYLE_BINDING_MAP_NAMESPACE}}}"
    if (
        root.tag != f"{namespace}documentHeadingStyleBindingMap"
        or dict(root.attrib) != {"version": "1"}
        or root.text is not None
        or root.tail is not None
    ):
        raise DocxSemanticsV3Error("Heading-style map root is not closed and canonical")
    output = []
    for node in root:
        if (
            node.tag != f"{namespace}binding"
            or set(node.attrib) != {"source_id", "level", "resolved_style_id", "visible_name"}
            or node.get("level") not in {str(index) for index in range(1, 10)}
            or node.text is not None
            or node.tail is not None
            or len(node)
        ):
            raise DocxSemanticsV3Error("Heading-style map record is not closed and canonical")
        output.append(
            HeadingStyleBindingV3(
                node.get("source_id"), int(node.get("level")), node.get("resolved_style_id"), node.get("visible_name")
            )
        )
    ordered = _validate(tuple(output))
    if ordered != tuple(output) or not ordered:
        raise DocxSemanticsV3Error("Heading-style map must contain canonically ordered bindings")
    return ordered


def prove_heading_style_registry(document: Any, bindings: tuple[HeadingStyleBindingV3, ...]) -> None:
    styles = _style_registry(document)
    for binding in _validate(bindings):
        style = _style(styles, binding.resolved_style_id)
        names = style.findall(qn("w:name"))
        if len(names) != 1 or names[0].get(qn("w:val")) != binding.visible_name:
            raise DocxSemanticsV3Error("Heading-style primary name differs from the recorded identity")


def prove_heading_paragraph(paragraph: Any, binding: HeadingStyleBindingV3) -> None:
    properties = paragraph.findall(qn("w:pPr"))
    references = properties[0].findall(qn("w:pStyle")) if len(properties) == 1 else []
    if (
        len(references) != 1
        or dict(references[0].attrib) != {qn("w:val"): binding.resolved_style_id}
        or len(references[0])
    ):
        raise DocxSemanticsV3Error("Heading target must retain its exact recorded paragraph style")
