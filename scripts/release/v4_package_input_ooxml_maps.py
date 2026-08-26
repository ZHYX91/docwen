from __future__ import annotations

import hashlib
import re
import uuid
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from typing import NoReturn
from xml.etree import ElementTree

RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
OFFICE_RELS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CUSTOM_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/customXml"
CAPTION_STYLE_NAMESPACE = "https://docwen.dev/schema/document-caption-style-binding-map/v1"
OCCURRENCE_NAMESPACE = "https://docwen.dev/schema/document-numbering-occurrence-map/v1"

_CUSTOM_XML_REL = f"{OFFICE_RELS}/customXml"
_CUSTOM_XML_PROPS_REL = f"{OFFICE_RELS}/customXmlProps"
_CUSTOM_XML_PROPS_TYPE = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"
_ITEM = re.compile(r"customXml/item([1-9][0-9]*)\.xml")
_STYLE_ID = re.compile(r"[A-Za-z][A-Za-z0-9]{0,252}")
_SHA256 = re.compile(r"[0-9a-f]{64}")
_STYLE_KEYS = ("figure_caption", "table_caption", "equation_caption", "code_block_caption")
_KEY_KIND = {
    "figure_caption": "figure",
    "table_caption": "table",
    "equation_caption": "equation",
    "code_block_caption": "code_block",
}


class V4OoxmlMapError(ValueError):
    """An owned custom-XML authority is absent, ambiguous, or forged."""


@dataclass(frozen=True)
class CaptionMaps:
    styles_by_kind: Mapping[str, tuple[str, str]]
    occurrence_tags: Mapping[tuple[int, int, str], str]


def _fail(code: str) -> NoReturn:
    raise V4OoxmlMapError(code)


def _xml(raw: bytes, code: str) -> ElementTree.Element:
    upper = raw.upper()
    if b"<!DOCTYPE" in upper or b"<!ENTITY" in upper:
        _fail(code)
    try:
        return ElementTree.fromstring(raw)
    except ElementTree.ParseError as exc:
        raise V4OoxmlMapError(code) from exc


def _namespace(root: ElementTree.Element) -> str | None:
    if not root.tag.startswith("{") or "}" not in root.tag:
        return None
    return root.tag[1 : root.tag.index("}")]


def _owned_items(parts: Mapping[str, bytes]) -> dict[str, tuple[int, str, bytes, ElementTree.Element]]:
    result: dict[str, tuple[int, str, bytes, ElementTree.Element]] = {}
    signals = tuple(value.encode() for value in (CAPTION_STYLE_NAMESPACE, OCCURRENCE_NAMESPACE))
    for name, raw in parts.items():
        match = _ITEM.fullmatch(name)
        if match is None:
            continue
        try:
            root = _xml(raw, "ooxml_custom_xml_invalid")
        except V4OoxmlMapError:
            if any(signal in raw for signal in signals):
                raise
            continue
        namespace = _namespace(root)
        if namespace not in {CAPTION_STYLE_NAMESPACE, OCCURRENCE_NAMESPACE}:
            continue
        if namespace in result:
            _fail("ooxml_owned_custom_xml_duplicate")
        result[namespace] = (int(match.group(1)), name, raw, root)
    canonical_signal_parts = {
        part for number, name, _raw, _root in result.values() for part in (name, f"customXml/itemProps{number}.xml")
    }
    for name, raw in parts.items():
        if name not in canonical_signal_parts and any(signal in raw for signal in signals):
            _fail("ooxml_owned_custom_xml_signal_in_noncanonical_part")
    return result


def _relationship_support(
    number: int,
    namespace: str,
    item_raw: bytes,
    parts: Mapping[str, bytes],
    document_relationships: ElementTree.Element,
    content_types: ElementTree.Element,
) -> None:
    item = f"customXml/item{number}.xml"
    props_name = f"customXml/itemProps{number}.xml"
    item_rels_name = f"customXml/_rels/item{number}.xml.rels"
    if props_name not in parts or item_rels_name not in parts:
        _fail("ooxml_custom_xml_support_missing")
    rel_tag = f"{{{RELS}}}Relationship"
    target = f"../{item}"
    matches = [entry for entry in document_relationships if entry.get("Target") == target]
    if (
        len(matches) != 1
        or matches[0].tag != rel_tag
        or set(matches[0].attrib) != {"Id", "Type", "Target"}
        or matches[0].get("Type") != _CUSTOM_XML_REL
        or matches[0].get("TargetMode") is not None
    ):
        _fail("ooxml_custom_xml_document_relationship_invalid")
    item_rels = _xml(parts[item_rels_name], "ooxml_custom_xml_item_relationship_invalid")
    children = list(item_rels)
    if (
        item_rels.tag != f"{{{RELS}}}Relationships"
        or item_rels.attrib
        or len(children) != 1
        or children[0].tag != rel_tag
        or children[0].attrib != {"Id": "rId1", "Type": _CUSTOM_XML_PROPS_REL, "Target": f"itemProps{number}.xml"}
    ):
        _fail("ooxml_custom_xml_item_relationship_invalid")
    props = _xml(parts[props_name], "ooxml_custom_xml_properties_invalid")
    props_ns = f"{{{CUSTOM_PROPS}}}"
    expected_uuid = (
        "{" + str(uuid.uuid5(uuid.NAMESPACE_URL, f"{namespace}\0{hashlib.sha256(item_raw).hexdigest()}")).upper() + "}"
    )
    refs = props.findall(f"{props_ns}schemaRefs/{props_ns}schemaRef")
    if (
        props.tag != f"{props_ns}datastoreItem"
        or props.attrib != {f"{props_ns}itemID": expected_uuid}
        or len(list(props)) != 1
        or len(refs) != 1
        or refs[0].attrib != {f"{props_ns}uri": namespace}
        or list(refs[0])
    ):
        _fail("ooxml_custom_xml_properties_invalid")
    item_part, props_part = f"/{item}", f"/{props_name}"
    item_types = [entry for entry in content_types if entry.get("PartName") == item_part]
    props_types = [entry for entry in content_types if entry.get("PartName") == props_part]
    if (
        len(item_types) != 1
        or item_types[0].tag != f"{{{CONTENT_TYPES}}}Override"
        or item_types[0].get("ContentType") != "application/xml"
        or len(props_types) != 1
        or props_types[0].tag != f"{{{CONTENT_TYPES}}}Override"
        or props_types[0].get("ContentType") != _CUSTOM_XML_PROPS_TYPE
    ):
        _fail("ooxml_custom_xml_content_type_invalid")


def _caption_styles(root: ElementTree.Element) -> dict[str, tuple[str, str]]:
    ns = f"{{{CAPTION_STYLE_NAMESPACE}}}"
    if (
        root.tag != f"{ns}documentCaptionStyleBindingMap"
        or root.attrib != {"version": "1"}
        or root.text is not None
        or root.tail is not None
    ):
        _fail("ooxml_caption_style_map_invalid")
    result: dict[str, tuple[str, str]] = {}
    for expected_key, item in zip(_STYLE_KEYS, list(root), strict=False):
        if (
            item.tag != f"{ns}binding"
            or tuple(item.attrib) != ("semantic_key", "resolved_style_id", "visible_name")
            or item.get("semantic_key") != expected_key
            or item.text is not None
            or item.tail is not None
            or list(item)
        ):
            _fail("ooxml_caption_style_binding_invalid")
        style_id, visible = item.get("resolved_style_id", ""), item.get("visible_name", "")
        if _STYLE_ID.fullmatch(style_id) is None or not visible or visible != visible.strip() or len(visible) > 255:
            _fail("ooxml_caption_style_binding_invalid")
        result[_KEY_KIND[expected_key]] = (style_id, visible)
    if len(list(root)) != 4 or len(result) != 4:
        _fail("ooxml_caption_style_inventory_invalid")
    if (
        len({value[0].casefold() for value in result.values()}) != 4
        or len({value[1].casefold() for value in result.values()}) != 4
    ):
        _fail("ooxml_caption_style_identity_duplicate")
    return result


def _range(target: Mapping[str, object]) -> tuple[int, int]:
    start, end = target.get("source_start"), target.get("source_end")
    if (
        not isinstance(start, int)
        or isinstance(start, bool)
        or not isinstance(end, int)
        or isinstance(end, bool)
        or start < 0
        or end <= start
    ):
        _fail("ooxml_occurrence_range_invalid")
    return start, end


def _derived_occurrence(
    target: Mapping[str, object], source_sha256: str, plan_sha256: str
) -> tuple[tuple[int, int, str], str, str]:
    start, end = _range(target)
    kind = str(target.get("kind"))
    preimage = f"docwen-numbering-occurrence-map-v1\0{source_sha256}\0{start}\0{end}\0{kind}\0false\0\0\0{plan_sha256}"
    digest = hashlib.sha256(preimage.encode()).hexdigest()
    return (start, end, kind), f"docwen-numbering-occurrence-v1:{digest[:32]}", digest


def _occurrences(
    root: ElementTree.Element,
    expected_targets: Sequence[Mapping[str, object]],
    source_sha256: str,
    plan_sha256: str,
) -> dict[tuple[int, int, str], str]:
    ns = f"{{{OCCURRENCE_NAMESPACE}}}"
    if (
        root.tag != f"{ns}documentNumberingOccurrenceMap"
        or tuple(root.attrib) != ("version", "plan_sha256")
        or root.get("version") != "1"
        or root.get("plan_sha256") != plan_sha256
        or root.text is not None
        or root.tail is not None
    ):
        _fail("ooxml_occurrence_map_invalid")
    expected = [_derived_occurrence(item, source_sha256, plan_sha256) for item in expected_targets]
    expected.sort(key=lambda item: item[0])
    children = list(root)
    attributes = (
        "tag",
        "source_sha256",
        "source_start",
        "source_end",
        "kind",
        "enabled",
        "target_id",
        "derived_number",
        "plan_sha256",
        "sha256",
    )
    if len(children) != len(expected):
        _fail("ooxml_occurrence_inventory_invalid")
    result: dict[tuple[int, int, str], str] = {}
    for item, (key, tag, digest) in zip(children, expected, strict=True):
        start, end, kind = key
        expected_values = {
            "tag": tag,
            "source_sha256": source_sha256,
            "source_start": str(start),
            "source_end": str(end),
            "kind": kind,
            "enabled": "false",
            "target_id": "",
            "derived_number": "",
            "plan_sha256": plan_sha256,
            "sha256": digest,
        }
        if (
            item.tag != f"{ns}occurrence"
            or tuple(item.attrib) != attributes
            or item.attrib != expected_values
            or item.text is not None
            or item.tail is not None
            or list(item)
        ):
            _fail("ooxml_occurrence_record_invalid")
        result[key] = tag
    return result


def inspect_caption_maps(
    parts: Mapping[str, bytes],
    document_relationships: ElementTree.Element,
    content_types: ElementTree.Element,
    targets: Sequence[Mapping[str, object]],
    planned_targets: Sequence[Mapping[str, object]],
    source_sha256: str,
    plan_sha256: str,
) -> CaptionMaps:
    """Authenticate caption styles and disabled ID-less occurrence identities."""

    if _SHA256.fullmatch(source_sha256) is None or _SHA256.fullmatch(plan_sha256) is None:
        _fail("ooxml_custom_xml_plan_identity_invalid")
    owned = _owned_items(parts)
    for namespace, (number, _name, raw, _root) in owned.items():
        _relationship_support(number, namespace, raw, parts, document_relationships, content_types)
    captions = [item for item in targets if item.get("kind") in _KEY_KIND.values()]
    style_item = owned.get(CAPTION_STYLE_NAMESPACE)
    if bool(captions) != bool(style_item):
        _fail("ooxml_caption_style_map_presence_invalid")
    styles = {} if style_item is None else _caption_styles(style_item[3])
    disabled_idless = [
        target
        for target, planned in zip(targets, planned_targets, strict=True)
        if target.get("kind") in _KEY_KIND.values()
        and target.get("target_id") is None
        and planned.get("enabled") is False
    ]
    occurrence_item = owned.get(OCCURRENCE_NAMESPACE)
    if bool(disabled_idless) != bool(occurrence_item):
        _fail("ooxml_occurrence_map_presence_invalid")
    occurrences = (
        {} if occurrence_item is None else _occurrences(occurrence_item[3], disabled_idless, source_sha256, plan_sha256)
    )
    return CaptionMaps(styles, occurrences)


__all__ = ["CaptionMaps", "V4OoxmlMapError", "inspect_caption_maps"]
