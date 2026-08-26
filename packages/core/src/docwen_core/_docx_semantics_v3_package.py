"""Deterministic custom-XML package writer and strict map reader for v3."""

from __future__ import annotations

import hashlib
import os
import re
import tempfile
import uuid
from itertools import pairwise
from pathlib import Path
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

import lxml.etree as etree

from docwen_core._docx_semantics_v3_fenced import FENCED_SOURCE_MAP_NAMESPACE
from docwen_core._docx_semantics_v3_fenced_map import fenced_source_map_xml, parse_fenced_source_map
from docwen_core._docx_semantics_v3_model import (
    ANCHOR_TOPOLOGY_MAP_NAMESPACE,
    CAPTION_STYLE_BINDING_MAP_NAMESPACE,
    REFERENCE_OCCURRENCE_MAP_NAMESPACE,
    SOFT_REFERENCE_MAP_NAMESPACE,
    TARGET_MAP_NAMESPACE,
    AnchorIdentityV3,
    DocxSemanticsV3Error,
    ReferenceOccurrenceIdentityV3,
    SoftReferenceIdentityV3,
    TargetIdentityV3,
    derive_anchor_identity_v3,
    derive_reference_occurrence_identity_v3,
    derive_soft_reference_identity_v3,
    derive_target_identity_v3,
)
from docwen_core._docx_semantics_v3_styles import caption_style_binding_map_xml, parse_caption_style_binding_map
from docwen_core._docx_semantics_v3_topology import anchor_topology_map_xml, parse_anchor_topology_map
from docwen_core.docx_citation_ooxml import (
    CITATION_ITEM_MAP_NAMESPACE,
    CITATION_OCCURRENCE_MAP_NAMESPACE,
    citation_item_map_xml,
    citation_occurrence_map_xml,
    parse_citation_item_map,
    parse_citation_occurrence_map,
)
from docwen_core.docx_numbering_occurrence import (
    NUMBERING_OCCURRENCE_MAP_NAMESPACE,
    numbering_occurrence_map_xml,
    parse_numbering_occurrence_map,
)

_RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships"
_CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types"
_CUSTOM_XML_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
_CUSTOM_XML_PROPS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps"
_CUSTOM_XML_PROPS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"
_CUSTOM_XML_PROPERTIES_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/customXml"
_XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
_ITEM_RE = re.compile(r"customXml/item(?P<number>[1-9][0-9]*)\.xml$")

_OWNED_MAP_NAMESPACES = frozenset(
    {
        ANCHOR_TOPOLOGY_MAP_NAMESPACE,
        CAPTION_STYLE_BINDING_MAP_NAMESPACE,
        FENCED_SOURCE_MAP_NAMESPACE,
        CITATION_ITEM_MAP_NAMESPACE,
        CITATION_OCCURRENCE_MAP_NAMESPACE,
        NUMBERING_OCCURRENCE_MAP_NAMESPACE,
        TARGET_MAP_NAMESPACE,
        SOFT_REFERENCE_MAP_NAMESPACE,
        REFERENCE_OCCURRENCE_MAP_NAMESPACE,
    }
)


def semantic_map_xml(targets: list[TargetIdentityV3], anchors: list[AnchorIdentityV3]) -> bytes:
    target_entries = "".join(
        (
            f'<target kind="{xml_attr(item.kind)}" source_id="{xml_attr(item.source_id)}" '
            f'bookmark_name="{item.bookmark_name}" sha256="{item.sha256}"/>'
        )
        for item in sorted(targets, key=lambda value: (value.kind, value.source_id))
    )
    anchor_entries = "".join(
        (
            f'<anchor block_kind="{xml_attr(item.block_kind)}" source_id="{xml_attr(item.source_id)}" '
            f'tag="{item.tag}" sha256="{item.sha256}"/>'
        )
        for item in sorted(anchors, key=lambda value: (value.block_kind, value.source_id))
    )
    root = (
        f'<documentSemanticMap xmlns="{TARGET_MAP_NAMESPACE}" version="1">'
        f"<targets>{target_entries}</targets><anchors>{anchor_entries}</anchors>"
        "</documentSemanticMap>"
    )
    return f"{_XML_DECLARATION}\n{root}\n".encode()


def soft_reference_map_xml(records: list[SoftReferenceIdentityV3]) -> bytes:
    entries = "".join(
        (
            f'<softReference tag="{item.tag}" source_sha256="{item.source_sha256}" '
            f'source_start="{item.source_start}" source_end="{item.source_end}" '
            f'authored_token="{xml_attr(item.authored_token)}" '
            f'cached_number="{xml_attr(item.cached_number)}"/>'
        )
        for item in sorted(records, key=lambda value: (value.source_start, value.source_end, value.tag))
    )
    root = (
        f'<documentSoftReferenceMap xmlns="{SOFT_REFERENCE_MAP_NAMESPACE}" version="1">'
        f"{entries}</documentSoftReferenceMap>"
    )
    return f"{_XML_DECLARATION}\n{root}\n".encode()


def reference_occurrence_map_xml(records: list[ReferenceOccurrenceIdentityV3]) -> bytes:
    entries = "".join(
        (
            f'<referenceOccurrence tag="{item.tag}" source_sha256="{item.source_sha256}" '
            f'source_start="{item.source_start}" source_end="{item.source_end}" '
            f'authored_token="{xml_attr(item.authored_token)}" '
            f'resolved_bookmark_name="{item.resolved_bookmark_name}" '
            f'cached_number="{xml_attr(item.cached_number)}"/>'
        )
        for item in sorted(records, key=lambda value: (value.source_start, value.source_end, value.tag))
    )
    root = (
        f'<documentReferenceOccurrenceMap xmlns="{REFERENCE_OCCURRENCE_MAP_NAMESPACE}" version="1">'
        f"{entries}</documentReferenceOccurrenceMap>"
    )
    return f"{_XML_DECLARATION}\n{root}\n".encode()


def xml_attr(value: str) -> str:
    return (
        value.replace("&", "&amp;")
        .replace('"', "&quot;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace("\r", "&#13;")
        .replace("\n", "&#10;")
        .replace("\t", "&#9;")
    )


def inject_custom_xml_parts(
    path: Path,
    parts: list[tuple[str, bytes]],
    *,
    allow_existing_owned: bool = False,
) -> None:
    if not path.is_file():
        raise DocxSemanticsV3Error("DOCX package does not exist")
    with ZipFile(path) as package:
        infos = package.infolist()
        if len({info.filename for info in infos}) != len(infos):
            raise DocxSemanticsV3Error("DOCX package contains duplicate ZIP members")
        existing = {info.filename: package.read(info.filename) for info in infos}
    if not allow_existing_owned:
        _preflight_no_owned_map(existing)
    used_item_ids = _custom_xml_item_ids(existing)
    rels_name = "word/_rels/document.xml.rels"
    types_name = "[Content_Types].xml"
    if rels_name not in existing or types_name not in existing:
        raise DocxSemanticsV3Error("DOCX package lacks document relationships or content types")
    parser = etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)
    rels_root = etree.fromstring(existing[rels_name], parser)
    types_root = etree.fromstring(existing[types_name], parser)
    _validate_relationships_root(rels_root, context="document")
    if types_root.tag != f"{{{_CONTENT_TYPES_NAMESPACE}}}Types":
        raise DocxSemanticsV3Error("content-types root is invalid")

    new_files: dict[str, bytes] = {}
    occupied_numbers = _occupied_item_numbers(existing)
    used_rids = {
        int(match.group(1))
        for relation in rels_root
        if (match := re.fullmatch(r"rId([1-9][0-9]*)", relation.get("Id", ""))) is not None
    }
    new_overrides: list[tuple[str, str]] = []
    for namespace, item_bytes in parts:
        item_number = _lowest_positive_not_in(occupied_numbers)
        occupied_numbers.add(item_number)
        relationship_number = _lowest_positive_not_in(used_rids)
        used_rids.add(relationship_number)
        item_name = f"customXml/item{item_number}.xml"
        props_name = f"customXml/itemProps{item_number}.xml"
        item_rels_name = f"customXml/_rels/item{item_number}.xml.rels"
        if any(name in existing or name in new_files for name in (item_name, props_name, item_rels_name)):
            raise DocxSemanticsV3Error("custom XML part allocation conflict")
        item_digest = hashlib.sha256(item_bytes).hexdigest()
        item_uuid = "{" + str(uuid.uuid5(uuid.NAMESPACE_URL, f"{namespace}\0{item_digest}")).upper() + "}"
        if item_uuid.casefold() in used_item_ids:
            raise DocxSemanticsV3Error("custom XML deterministic item UUID collides")
        used_item_ids.add(item_uuid.casefold())
        props_bytes = _canonical_item_properties(namespace, item_uuid)
        item_rels_root = (
            f'<Relationships xmlns="{_RELATIONSHIPS_NAMESPACE}">'
            f'<Relationship Id="rId1" Type="{_CUSTOM_XML_PROPS_REL}" '
            f'Target="itemProps{item_number}.xml"/></Relationships>'
        )
        new_files.update(
            {
                item_name: item_bytes,
                props_name: props_bytes,
                item_rels_name: f"{_XML_DECLARATION}\n{item_rels_root}\n".encode(),
            }
        )
        relation = etree.Element(f"{{{_RELATIONSHIPS_NAMESPACE}}}Relationship")
        relation.set("Id", f"rId{relationship_number}")
        relation.set("Type", _CUSTOM_XML_REL)
        relation.set("Target", f"../customXml/item{item_number}.xml")
        rels_root.append(relation)
        new_overrides.extend(
            (
                (f"/{item_name}", "application/xml"),
                (f"/{props_name}", _CUSTOM_XML_PROPS_CONTENT_TYPE),
            )
        )
    _append_content_type_overrides(types_root, new_overrides)
    existing[rels_name] = etree.tostring(rels_root, encoding="UTF-8", xml_declaration=True, standalone=True)
    existing[types_name] = etree.tostring(types_root, encoding="UTF-8", xml_declaration=True, standalone=True)
    existing.update(new_files)
    _rewrite_zip(path, infos, existing, tuple(new_files))


def read_owned_map_parts(package: ZipFile) -> dict[str, tuple[int, Any]]:
    parser = etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)
    output: dict[str, tuple[int, Any]] = {}
    names = package.namelist()
    if len(set(names)) != len(names):
        raise DocxSemanticsV3Error("DOCX package contains duplicate ZIP members")
    _verify_package_item_ids(package)
    _reject_malformed_owned_signals(
        {name: package.read(name) for name in names if name.startswith("customXml/")},
        parser,
    )
    for name in names:
        match = _ITEM_RE.fullmatch(name)
        if match is None:
            continue
        data = package.read(name)
        try:
            root = etree.fromstring(data, parser)
        except etree.XMLSyntaxError:
            continue
        namespace = etree.QName(root).namespace
        if namespace not in _OWNED_MAP_NAMESPACES:
            continue
        require_exact_xml_framing(data)
        _require_canonical_owned_map_bytes(namespace, root, data)
        if namespace in output:
            raise DocxSemanticsV3Error(f"duplicate custom XML map namespace: {namespace}")
        output[namespace] = (int(match.group("number")), root)
    return output


def verify_custom_xml_support(package: ZipFile, item_number: int, namespace: str) -> None:
    parser = etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)
    names = set(package.namelist())
    props_name = f"customXml/itemProps{item_number}.xml"
    item_rels_name = f"customXml/_rels/item{item_number}.xml.rels"
    if props_name not in names or item_rels_name not in names:
        raise DocxSemanticsV3Error("custom XML map lacks properties or item relationship")
    item_bytes = package.read(f"customXml/item{item_number}.xml")
    props_bytes = package.read(props_name)
    item_rels_bytes = package.read(item_rels_name)
    require_exact_xml_framing(props_bytes)
    require_exact_xml_framing(item_rels_bytes)
    item_rels = etree.fromstring(item_rels_bytes, parser)
    relationships = list(item_rels)
    if (
        item_rels.tag != f"{{{_RELATIONSHIPS_NAMESPACE}}}Relationships"
        or item_rels.attrib
        or item_rels.text is not None
        or item_rels.tail is not None
        or len(relationships) != 1
        or relationships[0].tag != f"{{{_RELATIONSHIPS_NAMESPACE}}}Relationship"
        or tuple(relationships[0].attrib) != ("Id", "Type", "Target")
        or relationships[0].get("Id") != "rId1"
        or relationships[0].get("Type") != _CUSTOM_XML_PROPS_REL
        or relationships[0].get("Target") != f"itemProps{item_number}.xml"
        or relationships[0].get("TargetMode") is not None
        or relationships[0].text is not None
        or relationships[0].tail is not None
        or len(relationships[0]) != 0
    ):
        raise DocxSemanticsV3Error("custom XML item relationship is not canonical")
    props = etree.fromstring(props_bytes, parser)
    props_children = list(props)
    schema_refs = props.findall(
        f"{{{_CUSTOM_XML_PROPERTIES_NAMESPACE}}}schemaRefs/{{{_CUSTOM_XML_PROPERTIES_NAMESPACE}}}schemaRef"
    )
    if (
        props.tag != f"{{{_CUSTOM_XML_PROPERTIES_NAMESPACE}}}datastoreItem"
        or tuple(props.attrib) != (f"{{{_CUSTOM_XML_PROPERTIES_NAMESPACE}}}itemID",)
        or props.text is not None
        or props.tail is not None
        or len(props_children) != 1
        or props_children[0].tag != f"{{{_CUSTOM_XML_PROPERTIES_NAMESPACE}}}schemaRefs"
        or props_children[0].attrib
        or props_children[0].text is not None
        or props_children[0].tail is not None
        or len(props_children[0]) != 1
        or len(schema_refs) != 1
        or tuple(schema_refs[0].attrib) != (f"{{{_CUSTOM_XML_PROPERTIES_NAMESPACE}}}uri",)
        or schema_refs[0].text is not None
        or schema_refs[0].tail is not None
        or len(schema_refs[0]) != 0
        or schema_refs[0].get(f"{{{_CUSTOM_XML_PROPERTIES_NAMESPACE}}}uri") != namespace
    ):
        raise DocxSemanticsV3Error("custom XML properties do not bind the expected schema")
    expected_uuid = (
        "{"
        + str(
            uuid.uuid5(
                uuid.NAMESPACE_URL,
                f"{namespace}\0{hashlib.sha256(item_bytes).hexdigest()}",
            )
        ).upper()
        + "}"
    )
    if props.get(f"{{{_CUSTOM_XML_PROPERTIES_NAMESPACE}}}itemID") != expected_uuid:
        raise DocxSemanticsV3Error("custom XML item UUID does not match its deterministic identity")
    if props_bytes != _canonical_item_properties(namespace, expected_uuid):
        raise DocxSemanticsV3Error("custom XML properties bytes are not canonical")
    _verify_document_relationship(package, item_number, parser)
    _verify_content_types(package, item_number, parser)


def parse_semantic_map(root: Any) -> tuple[list[TargetIdentityV3], list[AnchorIdentityV3]]:
    namespace = f"{{{TARGET_MAP_NAMESPACE}}}"
    if root.tag != f"{namespace}documentSemanticMap" or tuple(root.attrib.items()) != (("version", "1"),):
        raise DocxSemanticsV3Error("semantic-map root is not canonical")
    children = list(root)
    if [item.tag for item in children] != [f"{namespace}targets", f"{namespace}anchors"]:
        raise DocxSemanticsV3Error("semantic-map container order is invalid")
    targets = [_parse_target(item, namespace) for item in children[0]]
    anchors = [_parse_anchor(item, namespace) for item in children[1]]
    if len({item.source_id for item in [*targets, *anchors]}) != len(targets) + len(anchors):
        raise DocxSemanticsV3Error("semantic and ordinary anchors do not share a unique namespace")
    if targets != sorted(targets, key=lambda item: (item.kind, item.source_id)) or anchors != sorted(
        anchors, key=lambda item: (item.block_kind, item.source_id)
    ):
        raise DocxSemanticsV3Error("semantic-map records are not canonically ordered")
    return targets, anchors


def parse_soft_reference_map(root: Any) -> list[SoftReferenceIdentityV3]:
    namespace = f"{{{SOFT_REFERENCE_MAP_NAMESPACE}}}"
    if root.tag != f"{namespace}documentSoftReferenceMap" or tuple(root.attrib.items()) != (("version", "1"),):
        raise DocxSemanticsV3Error("soft-reference root is not canonical")
    records = [_parse_soft_reference(item, namespace) for item in root]
    if len({item.tag for item in records}) != len(records):
        raise DocxSemanticsV3Error("soft-reference tags are not unique")
    if records != sorted(records, key=lambda item: (item.source_start, item.source_end, item.tag)):
        raise DocxSemanticsV3Error("soft-reference records are not canonically ordered")
    for previous, current in pairwise(records):
        if current.source_start < previous.source_end:
            raise DocxSemanticsV3Error("soft-reference source ranges overlap")
    return records


def parse_reference_occurrence_map(root: Any) -> list[ReferenceOccurrenceIdentityV3]:
    namespace = f"{{{REFERENCE_OCCURRENCE_MAP_NAMESPACE}}}"
    if root.tag != f"{namespace}documentReferenceOccurrenceMap" or tuple(root.attrib.items()) != (("version", "1"),):
        raise DocxSemanticsV3Error("reference-occurrence root is not canonical")
    if root.text is not None or root.tail is not None:
        raise DocxSemanticsV3Error("reference-occurrence root contains text")
    records = [_parse_reference_occurrence(item, namespace) for item in root]
    if len({item.tag for item in records}) != len(records):
        raise DocxSemanticsV3Error("reference-occurrence tags are not unique")
    if records != sorted(records, key=lambda item: (item.source_start, item.source_end, item.tag)):
        raise DocxSemanticsV3Error("reference-occurrence records are not canonically ordered")
    for previous, current in pairwise(records):
        if current.source_start < previous.source_end:
            raise DocxSemanticsV3Error("reference-occurrence source ranges overlap")
    return records


def require_exact_xml_framing(data: bytes) -> None:
    if (
        data.startswith(b"\xef\xbb\xbf")
        or b"\r" in data
        or not data.startswith((_XML_DECLARATION + "\n").encode())
        or not data.endswith(b"\n")
        or data.count(b"\n") != 2
    ):
        raise DocxSemanticsV3Error("owned custom XML part has non-canonical byte framing")


def _occupied_item_numbers(existing: dict[str, bytes]) -> set[int]:
    occupied: set[int] = set()
    for name in existing:
        for pattern in (
            r"customXml/item([1-9][0-9]*)\.xml",
            r"customXml/itemProps([1-9][0-9]*)\.xml",
            r"customXml/_rels/item([1-9][0-9]*)\.xml\.rels",
        ):
            match = re.fullmatch(pattern, name)
            if match is not None:
                occupied.add(int(match.group(1)))
    return occupied


def _custom_xml_item_ids(existing: dict[str, bytes]) -> set[str]:
    parser = etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)
    output: set[str] = set()
    for name, data in existing.items():
        if not name.startswith("customXml/") or not name.endswith(".xml"):
            continue
        canonical_props = re.fullmatch(r"customXml/itemProps[1-9][0-9]*\.xml", name) is not None
        try:
            root = etree.fromstring(data, parser)
        except etree.XMLSyntaxError as exc:
            if canonical_props:
                raise DocxSemanticsV3Error("custom XML properties part is malformed") from exc
            continue
        values = [
            value
            for element in root.iter()
            if (value := element.get(f"{{{_CUSTOM_XML_PROPERTIES_NAMESPACE}}}itemID")) is not None
        ]
        if canonical_props and len(values) != 1:
            raise DocxSemanticsV3Error("custom XML item UUID is missing or duplicated in its part")
        for value in values:
            if (
                re.fullmatch(
                    r"\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}",
                    value,
                )
                is None
            ):
                raise DocxSemanticsV3Error("custom XML item UUID is missing or malformed")
            key = value.casefold()
            if key in output:
                raise DocxSemanticsV3Error("custom XML item UUID is duplicated")
            output.add(key)
    return output


def _verify_package_item_ids(package: ZipFile) -> None:
    data = {
        name: package.read(name)
        for name in package.namelist()
        if name.startswith("customXml/") and name.endswith(".xml")
    }
    _custom_xml_item_ids(data)


def _append_content_type_overrides(types_root: Any, new_overrides: list[tuple[str, str]]) -> None:
    existing_names = {element.get("PartName") for element in types_root if element.get("PartName") is not None}
    for part_name, content_type in sorted(new_overrides):
        if part_name in existing_names:
            raise DocxSemanticsV3Error("custom XML content-type override conflicts")
        override = etree.Element(f"{{{_CONTENT_TYPES_NAMESPACE}}}Override")
        override.set("PartName", part_name)
        override.set("ContentType", content_type)
        types_root.append(override)


def _preflight_no_owned_map(existing: dict[str, bytes]) -> None:
    parser = etree.XMLParser(resolve_entities=False, no_network=True)
    _reject_malformed_owned_signals(existing, parser)
    for name, data in existing.items():
        if _ITEM_RE.fullmatch(name) is None:
            continue
        try:
            root = etree.fromstring(data, parser)
        except etree.XMLSyntaxError:
            continue
        if etree.QName(root).namespace in _OWNED_MAP_NAMESPACES:
            raise DocxSemanticsV3Error("DOCX package already contains an owned v3 semantic map")


def _reject_malformed_owned_signals(existing: dict[str, bytes], parser: Any) -> None:
    signaled_numbers: set[int] = set()
    for name, data in existing.items():
        props_match = re.fullmatch(r"customXml/itemProps([1-9][0-9]*)\.xml", name)
        if props_match is None:
            continue
        literal_signal = any(namespace.encode("utf-8") in data for namespace in _OWNED_MAP_NAMESPACES)
        try:
            props = etree.fromstring(data, parser)
        except etree.XMLSyntaxError as exc:
            if literal_signal:
                raise DocxSemanticsV3Error("owned custom XML properties part is malformed") from exc
            continue
        schema_uris = {
            item.get(f"{{{_CUSTOM_XML_PROPERTIES_NAMESPACE}}}uri")
            for item in props.iter(f"{{{_CUSTOM_XML_PROPERTIES_NAMESPACE}}}schemaRef")
        }
        if literal_signal or schema_uris.intersection(_OWNED_MAP_NAMESPACES):
            signaled_numbers.add(int(props_match.group(1)))
    for number in signaled_numbers:
        item_name = f"customXml/item{number}.xml"
        data = existing.get(item_name)
        if data is None:
            raise DocxSemanticsV3Error("owned custom XML signal has no item part")
        try:
            root = etree.fromstring(data, parser)
        except etree.XMLSyntaxError as exc:
            raise DocxSemanticsV3Error("owned custom XML item part is malformed") from exc
        if etree.QName(root).namespace not in _OWNED_MAP_NAMESPACES:
            raise DocxSemanticsV3Error("owned custom XML properties point to a non-owned root")


def _require_canonical_owned_map_bytes(namespace: str, root: Any, data: bytes) -> None:
    if namespace == TARGET_MAP_NAMESPACE:
        targets, anchors = parse_semantic_map(root)
        expected = semantic_map_xml(targets, anchors)
    elif namespace == ANCHOR_TOPOLOGY_MAP_NAMESPACE:
        expected = anchor_topology_map_xml(parse_anchor_topology_map(root))
    elif namespace == SOFT_REFERENCE_MAP_NAMESPACE:
        expected = soft_reference_map_xml(parse_soft_reference_map(root))
    elif namespace == REFERENCE_OCCURRENCE_MAP_NAMESPACE:
        expected = reference_occurrence_map_xml(parse_reference_occurrence_map(root))
    elif namespace == CAPTION_STYLE_BINDING_MAP_NAMESPACE:
        expected = caption_style_binding_map_xml(parse_caption_style_binding_map(root))
    elif namespace == FENCED_SOURCE_MAP_NAMESPACE:
        expected = fenced_source_map_xml(parse_fenced_source_map(root))
    elif namespace == NUMBERING_OCCURRENCE_MAP_NAMESPACE:
        expected = numbering_occurrence_map_xml(parse_numbering_occurrence_map(root))
    elif namespace == CITATION_ITEM_MAP_NAMESPACE:
        expected = citation_item_map_xml(parse_citation_item_map(root))
    elif namespace == CITATION_OCCURRENCE_MAP_NAMESPACE:
        expected = citation_occurrence_map_xml(parse_citation_occurrence_map(root))
    else:  # pragma: no cover - caller owns the closed namespace set
        raise DocxSemanticsV3Error("unknown owned custom XML namespace")
    if data != expected:
        raise DocxSemanticsV3Error("owned custom XML bytes are not canonical")


def _lowest_positive_not_in(values: set[int]) -> int:
    candidate = 1
    while candidate in values:
        candidate += 1
    return candidate


def _rewrite_zip(
    path: Path,
    original_infos: list[ZipInfo],
    data: dict[str, bytes],
    new_names: tuple[str, ...],
) -> None:
    descriptor, temporary_name = tempfile.mkstemp(prefix=f".{path.name}.", suffix=".tmp", dir=path.parent)
    os.close(descriptor)
    temporary = Path(temporary_name)
    try:
        with ZipFile(temporary, "w") as output:
            original_names = {info.filename for info in original_infos}
            for info in original_infos:
                output.writestr(info, data[info.filename])
            for name in new_names:
                if name in original_names:
                    continue
                info = ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                info.external_attr = 0o600 << 16
                output.writestr(info, data[name])
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def _verify_document_relationship(package: ZipFile, item_number: int, parser: Any) -> None:
    rels = etree.fromstring(package.read("word/_rels/document.xml.rels"), parser)
    _validate_relationships_root(rels, context="document")
    target = f"../customXml/item{item_number}.xml"
    matching = [item for item in rels if item.get("Target") == target]
    if len(matching) != 1:
        raise DocxSemanticsV3Error("document does not own exactly one custom XML relationship")
    if (
        matching[0].tag != f"{{{_RELATIONSHIPS_NAMESPACE}}}Relationship"
        or tuple(matching[0].attrib) != ("Id", "Type", "Target")
        or matching[0].get("Type") != _CUSTOM_XML_REL
        or matching[0].text is not None
        or matching[0].tail is not None
        or len(matching[0]) != 0
    ):
        raise DocxSemanticsV3Error("owned document relationship is not closed and canonical")


def _verify_content_types(package: ZipFile, item_number: int, parser: Any) -> None:
    types = etree.fromstring(package.read("[Content_Types].xml"), parser)
    if types.tag != f"{{{_CONTENT_TYPES_NAMESPACE}}}Types" or types.attrib:
        raise DocxSemanticsV3Error("content-types root is invalid")
    item_part = f"/customXml/item{item_number}.xml"
    props_part = f"/customXml/itemProps{item_number}.xml"
    items = [item for item in types if item.get("PartName") == item_part]
    props = [item for item in types if item.get("PartName") == props_part]
    if (
        len(items) != 1
        or items[0].tag != f"{{{_CONTENT_TYPES_NAMESPACE}}}Override"
        or items[0].get("ContentType") != "application/xml"
        or tuple(items[0].attrib) != ("PartName", "ContentType")
        or items[0].text is not None
        or items[0].tail is not None
        or len(items[0]) != 0
        or len(props) != 1
        or props[0].tag != f"{{{_CONTENT_TYPES_NAMESPACE}}}Override"
        or props[0].get("ContentType") != _CUSTOM_XML_PROPS_CONTENT_TYPE
        or tuple(props[0].attrib) != ("PartName", "ContentType")
        or props[0].text is not None
        or props[0].tail is not None
        or len(props[0]) != 0
    ):
        raise DocxSemanticsV3Error("custom XML content types are missing or wrong")


def _validate_relationships_root(root: Any, *, context: str) -> None:
    relationship_tag = f"{{{_RELATIONSHIPS_NAMESPACE}}}Relationship"
    if root.tag != f"{{{_RELATIONSHIPS_NAMESPACE}}}Relationships" or root.attrib:
        raise DocxSemanticsV3Error(f"{context} relationships root is invalid")
    identifiers: list[str] = []
    for relation in root:
        identifier = relation.get("Id")
        if (
            relation.tag != relationship_tag
            or identifier is None
            or relation.get("Type") is None
            or relation.get("Target") is None
            or not set(relation.attrib).issubset({"Id", "Type", "Target", "TargetMode"})
            or relation.text is not None
            or (relation.tail is not None and relation.tail.strip())
            or len(relation) != 0
        ):
            raise DocxSemanticsV3Error(f"{context} relationship topology is invalid")
        identifiers.append(identifier)
    if len(set(identifiers)) != len(identifiers):
        raise DocxSemanticsV3Error(f"{context} relationship Id is duplicated")


def _canonical_item_properties(namespace: str, item_uuid: str) -> bytes:
    root = (
        f'<ds:datastoreItem xmlns:ds="{_CUSTOM_XML_PROPERTIES_NAMESPACE}" ds:itemID="{item_uuid}">'
        f'<ds:schemaRefs><ds:schemaRef ds:uri="{namespace}"/></ds:schemaRefs></ds:datastoreItem>'
    )
    return f"{_XML_DECLARATION}\n{root}\n".encode()


def _parse_target(item: Any, namespace: str) -> TargetIdentityV3:
    if item.tag != f"{namespace}target" or tuple(item.attrib) != ("kind", "source_id", "bookmark_name", "sha256"):
        raise DocxSemanticsV3Error("semantic target record is not closed and canonical")
    identity = derive_target_identity_v3(item.get("kind"), item.get("source_id"))  # type: ignore[arg-type]
    if item.get("bookmark_name") != identity.bookmark_name or item.get("sha256") != identity.sha256:
        raise DocxSemanticsV3Error("semantic target identity/hash does not recompute")
    return identity


def _parse_anchor(item: Any, namespace: str) -> AnchorIdentityV3:
    if item.tag != f"{namespace}anchor" or tuple(item.attrib) != ("block_kind", "source_id", "tag", "sha256"):
        raise DocxSemanticsV3Error("ordinary-anchor record is not closed and canonical")
    identity = derive_anchor_identity_v3(item.get("block_kind"), item.get("source_id"))
    if item.get("tag") != identity.tag or item.get("sha256") != identity.sha256:
        raise DocxSemanticsV3Error("ordinary-anchor identity/hash does not recompute")
    return identity


def _parse_soft_reference(item: Any, namespace: str) -> SoftReferenceIdentityV3:
    if item.tag != f"{namespace}softReference" or tuple(item.attrib) != (
        "tag",
        "source_sha256",
        "source_start",
        "source_end",
        "authored_token",
        "cached_number",
    ):
        raise DocxSemanticsV3Error("soft-reference record is not closed and canonical")
    try:
        identity = derive_soft_reference_identity_v3(
            source_sha256=item.get("source_sha256"),
            source_start=int(item.get("source_start")),
            source_end=int(item.get("source_end")),
            authored_token=item.get("authored_token"),
            cached_number=item.get("cached_number"),
        )
    except (TypeError, ValueError) as exc:
        raise DocxSemanticsV3Error("soft-reference record has invalid scalar values") from exc
    if item.get("tag") != identity.tag:
        raise DocxSemanticsV3Error("soft-reference digest does not recompute")
    return identity


def _parse_reference_occurrence(
    item: Any,
    namespace: str,
) -> ReferenceOccurrenceIdentityV3:
    if (
        item.tag != f"{namespace}referenceOccurrence"
        or tuple(item.attrib)
        != (
            "tag",
            "source_sha256",
            "source_start",
            "source_end",
            "authored_token",
            "resolved_bookmark_name",
            "cached_number",
        )
        or item.text is not None
        or item.tail is not None
        or len(item) != 0
    ):
        raise DocxSemanticsV3Error("reference-occurrence record is not closed and canonical")
    try:
        identity = derive_reference_occurrence_identity_v3(
            source_sha256=item.get("source_sha256"),
            source_start=int(item.get("source_start")),
            source_end=int(item.get("source_end")),
            authored_token=item.get("authored_token"),
            resolved_bookmark_name=item.get("resolved_bookmark_name"),
            cached_number=item.get("cached_number"),
        )
    except (TypeError, ValueError) as exc:
        raise DocxSemanticsV3Error("reference-occurrence record has invalid scalar values") from exc
    if item.get("tag") != identity.tag:
        raise DocxSemanticsV3Error("reference-occurrence digest does not recompute")
    return identity
