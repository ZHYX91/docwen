"""Resolve the main document's numbering part by its OPC relationship."""

from __future__ import annotations

import posixpath
from dataclasses import dataclass
from urllib.parse import unquote, urlsplit

from lxml import etree

WML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
NUMBERING_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
NUMBERING_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
_RELS_PART = "word/_rels/document.xml.rels"
_TYPES_PART = "[Content_Types].xml"


class NumberingPartError(ValueError):
    """The package cannot identify one safe numbering part."""


@dataclass(frozen=True, slots=True)
class NumberingPart:
    """The exact relationship-owned package member and parsed XML."""

    name: str
    element: etree._Element


def _parse(raw: bytes, expected: str, description: str) -> etree._Element:
    try:
        root = etree.fromstring(raw, etree.XMLParser(resolve_entities=False, no_network=True))
    except etree.XMLSyntaxError as exc:
        raise NumberingPartError(f"Invalid {description} XML") from exc
    if root.tag != expected:
        raise NumberingPartError(f"Invalid {description} root")
    return root


def _target_part(target: str) -> str:
    try:
        parsed = urlsplit(target)
    except ValueError as exc:
        raise NumberingPartError("Numbering relationship target is not a valid part URI") from exc
    if (
        not target
        or parsed.scheme
        or parsed.netloc
        or parsed.query
        or parsed.fragment
        or "\\" in target
        or any(ord(character) <= 32 for character in target)
    ):
        raise NumberingPartError("Numbering relationship target is not an internal part URI")
    # ZIP members use OPC part names. Do not decode their names or guess an
    # alternative member; encoded separators and traversal are ambiguous.
    for segment in target.split("/"):
        decoded = unquote(segment)
        if "/" in decoded or "\\" in decoded or (decoded in {".", ".."} and decoded != segment):
            raise NumberingPartError("Numbering relationship target contains an ambiguous URI segment")
    name = posixpath.normpath(target.lstrip("/") if target.startswith("/") else posixpath.join("word", target))
    if name in {"", ".", ".."} or name.startswith("../") or target.endswith("/"):
        raise NumberingPartError("Numbering relationship target escapes the package")
    return name


def _content_type(types: etree._Element, name: str) -> str | None:
    overrides = [
        item for item in types.findall(f"{{{CONTENT_TYPES_NS}}}Override") if item.get("PartName") == f"/{name}"
    ]
    if len(overrides) > 1:
        raise NumberingPartError("Numbering part has duplicate content-type overrides")
    if overrides:
        return overrides[0].get("ContentType")
    extension = posixpath.basename(name).rsplit(".", 1)[-1]
    defaults = [
        item
        for item in types.findall(f"{{{CONTENT_TYPES_NS}}}Default")
        if item.get("Extension", "").casefold() == extension.casefold()
    ]
    if len(defaults) > 1:
        raise NumberingPartError("Numbering part has ambiguous default content types")
    return defaults[0].get("ContentType") if defaults else None


def _xml(root: etree._Element) -> bytes:
    return etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)


def open_numbering_part(parts: dict[str, bytes], *, create: bool = False) -> NumberingPart | None:
    """Read the sole related numbering part, optionally creating an absent one.

    Existing relationship and content-type bytes are left untouched. Creation
    modifies only this in-memory package mapping; callers publish atomically
    after their semantic checks. Unrelated or orphan package members are never
    reused as an empty numbering part.
    """
    if _RELS_PART not in parts or _TYPES_PART not in parts:
        raise NumberingPartError("DOCX lacks required OPC support parts")
    rels = _parse(parts[_RELS_PART], f"{{{RELS_NS}}}Relationships", "document relationships")
    types = _parse(parts[_TYPES_PART], f"{{{CONTENT_TYPES_NS}}}Types", "content types")
    relationships = rels.findall(f"{{{RELS_NS}}}Relationship")
    owned = [item for item in relationships if item.get("Type") == NUMBERING_REL_TYPE]
    if len(owned) > 1:
        raise NumberingPartError("DOCX has multiple numbering relationships")
    if owned:
        relation = owned[0]
        identity = relation.get("Id")
        if not identity or sum(item.get("Id") == identity for item in relationships) != 1:
            raise NumberingPartError("Numbering relationship ID is absent or ambiguous")
        if relation.get("TargetMode", "Internal") != "Internal":
            raise NumberingPartError("Numbering relationship cannot be external")
        name = _target_part(relation.get("Target", ""))
        if name not in parts:
            raise NumberingPartError("Numbering relationship references a missing part")
        if _content_type(types, name) != NUMBERING_CONTENT_TYPE:
            raise NumberingPartError("Numbering part lacks its exact content type")
        return NumberingPart(name, _parse(parts[name], f"{{{WML_NS}}}numbering", "numbering part"))

    if any(item.get("ContentType") == NUMBERING_CONTENT_TYPE for item in types):
        raise NumberingPartError("Numbering part lacks its exact relationship")
    if not create:
        return None

    reserved = set(parts)
    reserved.update(item.get("PartName", "").lstrip("/") for item in types)
    for relation in relationships:
        if relation.get("TargetMode", "Internal") == "Internal":
            try:
                reserved.add(_target_part(relation.get("Target", "")))
            except (NumberingPartError, ValueError):
                # An unrelated malformed relationship is not ours to repair.
                continue
    name = "word/numbering.xml"
    suffix = 1
    while name in reserved:
        name = f"word/numbering{suffix}.xml"
        suffix += 1
    used_ids = {item.get("Id") for item in relationships}
    relation_index = 1
    while f"rId{relation_index}" in used_ids:
        relation_index += 1
    relation = etree.SubElement(rels, f"{{{RELS_NS}}}Relationship")
    relation.set("Id", f"rId{relation_index}")
    relation.set("Type", NUMBERING_REL_TYPE)
    relation.set("Target", posixpath.relpath(name, "word"))
    override = etree.SubElement(types, f"{{{CONTENT_TYPES_NS}}}Override")
    override.set("PartName", f"/{name}")
    override.set("ContentType", NUMBERING_CONTENT_TYPE)
    root = etree.Element(f"{{{WML_NS}}}numbering", nsmap={"w": WML_NS})
    parts[_RELS_PART] = _xml(rels)
    parts[_TYPES_PART] = _xml(types)
    parts[name] = _xml(root)
    return NumberingPart(name, root)
