"""Resolve or create numbering through the shared DOCX part identities."""

from __future__ import annotations

import posixpath

from lxml import etree

from docwen_core.docx_package_parts import (
    CONTENT_TYPES_NS,
    NUMBERING,
    RELS_NS,
    WML_NS,
    DocxPackageParts,
    DocxPartError,
    RelatedPart,
    relationship_part_name,
    resolve_part_target,
)


class NumberingPartError(ValueError):
    """The package cannot identify one safe numbering part."""


def _xml(root: etree._Element) -> bytes:
    return etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)


def open_numbering_part(parts: dict[str, bytes], *, create: bool = False) -> RelatedPart | None:
    """Read the relationship-owned part; create only when no numbering exists.

    Existing support bytes remain unchanged. Creation updates this request's
    mapping only, and never adopts an orphan or unrelated conventional member.
    """
    try:
        index = DocxPackageParts(parts)
        document = index.document()
        existing = index.one(document.name, NUMBERING)
        if existing is not None:
            return existing
        types = etree.fromstring(parts["[Content_Types].xml"])
        if any(item.get("ContentType") == NUMBERING.content_type for item in types):
            raise NumberingPartError("Numbering part lacks its exact relationship")
        if not create:
            return None
        rels = index.relationships(document.name)
        reserved = set(parts)
        reserved.update(item.get("PartName", "").lstrip("/") for item in types)
        for relation in rels:
            if relation.get("TargetMode", "Internal") == "Internal":
                try:
                    reserved.add(resolve_part_target(document.name, relation.get("Target", "")))
                except DocxPartError:
                    # Unrelated malformed targets are not ours to repair.
                    continue
        directory = posixpath.dirname(document.name)
        name = posixpath.join(directory, "numbering.xml")
        suffix = 1
        while name in reserved:
            name = posixpath.join(directory, f"numbering{suffix}.xml")
            suffix += 1
        used_ids = {item.get("Id") for item in rels}
        relation_index = 1
        while f"rId{relation_index}" in used_ids:
            relation_index += 1
        relation = etree.SubElement(rels, f"{{{RELS_NS}}}Relationship")
        relation.set("Id", f"rId{relation_index}")
        relation.set("Type", NUMBERING.relationship_type)
        relation.set("Target", posixpath.relpath(name, directory or "."))
        override = etree.SubElement(types, f"{{{CONTENT_TYPES_NS}}}Override")
        override.set("PartName", f"/{name}")
        override.set("ContentType", NUMBERING.content_type)
        root = etree.Element(f"{{{WML_NS}}}numbering", nsmap={"w": WML_NS})
        parts[relationship_part_name(document.name)] = _xml(rels)
        parts["[Content_Types].xml"] = _xml(types)
        parts[name] = _xml(root)
        return RelatedPart(name, root)
    except DocxPartError as exc:
        raise NumberingPartError(str(exc)) from exc
