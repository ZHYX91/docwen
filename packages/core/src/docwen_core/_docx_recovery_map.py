"""Exact-neutral DOCX recovery map: whole-package physical projection authority.

The recovery map is an optional request-owned custom-XML trio that binds three raw
pointers (neutral raw, plan raw, authored source) and a whole-package physical
projection digest to the same resolved-v4 DOCX package.  The projection digest is
computed before the recovery map trio is injected and, at recovery time, recomputed
under an identical exclusion rule so that the map can never be a self-hash.  A
package without the map is generic proof-only extraction; a package with a map that
fails projection proof is a hard failure.
"""

from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import lxml.etree as etree

from docwen_core._docx_semantics_v3_package import (
    DocxSemanticsV3Error,
    inject_custom_xml_parts,
    xml_attr,
)

RESOLVED_V4_RECOVERY_MAP_NAMESPACE = "https://docwen.dev/schema/resolved-v4-recovery-map/v1"
RECOVERY_PROJECTION_ALGORITHM = "docwen-ooxml-physical-v1"
RECOVERY_POINTER_ROLES = frozenset({"neutral_raw", "plan_raw", "authored_source"})
_XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
_DOC_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


@dataclass(frozen=True, slots=True)
class ResolvedV4RecoveryPointer:
    role: str
    relative_path: str
    bytes: int
    sha256: str


@dataclass(frozen=True, slots=True)
class ResolvedV4RecoveryBibliography:
    owner: str
    placeholder: str
    media_type: str


@dataclass(frozen=True, slots=True)
class ResolvedV4RecoveryMap:
    source_sha256: str
    plan_sha256: str
    physical_sha256: str
    pointers: tuple[ResolvedV4RecoveryPointer, ...]
    bibliography: ResolvedV4RecoveryBibliography
    item_number: int | None = None


@dataclass(frozen=True, slots=True)
class ResolvedV4RecoveryInput:
    """Caller-supplied raw bytes and staging names bound by one recovery map."""

    neutral_raw: bytes
    plan_raw: bytes
    authored_source: bytes
    neutral_name: str
    plan_name: str
    authored_name: str
    bibliography_owner: str
    bibliography_placeholder: str
    bibliography_media_type: str


def recovery_map_xml(value: ResolvedV4RecoveryMap) -> bytes:
    pointers = "".join(
        (
            f'<pointer role="{xml_attr(item.role)}" relative_path="{xml_attr(item.relative_path)}" '
            f'bytes="{item.bytes}" sha256="{item.sha256}"/>'
        )
        for item in sorted(value.pointers, key=lambda entry: (entry.role, entry.relative_path))
    )
    bibliography = (
        f'<bibliography owner="{xml_attr(value.bibliography.owner)}" '
        f'placeholder="{xml_attr(value.bibliography.placeholder)}" '
        f'media_type="{xml_attr(value.bibliography.media_type)}"/>'
    )
    root = (
        f'<resolvedV4RecoveryMap xmlns="{RESOLVED_V4_RECOVERY_MAP_NAMESPACE}" '
        f'version="1" source_sha256="{value.source_sha256}" '
        f'plan_sha256="{value.plan_sha256}" recovery_sha256="{_content_digest(value)}">'
        f'<pointers>{pointers}</pointers><projection version="1" '
        f'algorithm="{RECOVERY_PROJECTION_ALGORITHM}" physical_sha256="{value.physical_sha256}"/>'
        f"{bibliography}</resolvedV4RecoveryMap>"
    )
    return f"{_XML_DECLARATION}\n{root}\n".encode()


def _content_digest(value: ResolvedV4RecoveryMap) -> str:
    pointers = "".join(
        (
            f'<pointer role="{xml_attr(item.role)}" relative_path="{xml_attr(item.relative_path)}" '
            f'bytes="{item.bytes}" sha256="{item.sha256}"/>'
        )
        for item in sorted(value.pointers, key=lambda entry: (entry.role, entry.relative_path))
    )
    bibliography = (
        f'<bibliography owner="{xml_attr(value.bibliography.owner)}" '
        f'placeholder="{xml_attr(value.bibliography.placeholder)}" '
        f'media_type="{xml_attr(value.bibliography.media_type)}"/>'
    )
    projection = (
        f'<projection version="1" algorithm="{RECOVERY_PROJECTION_ALGORITHM}" '
        f'physical_sha256="{value.physical_sha256}"/>'
    )
    return hashlib.sha256(f"<pointers>{pointers}</pointers>{projection}{bibliography}".encode()).hexdigest()


def parse_recovery_map(root: Any) -> ResolvedV4RecoveryMap:
    if root.tag != f"{{{RESOLVED_V4_RECOVERY_MAP_NAMESPACE}}}resolvedV4RecoveryMap":
        raise DocxSemanticsV3Error("recovery map root element is invalid")
    expected_attributes = (
        "version",
        "source_sha256",
        "plan_sha256",
        "recovery_sha256",
    )
    if tuple(root.attrib) != expected_attributes:
        raise DocxSemanticsV3Error("recovery map root attribute order is invalid")
    if root.get("version") != "1":
        raise DocxSemanticsV3Error("recovery map version is not 1")
    source_sha256 = root.get("source_sha256") or ""
    plan_sha256 = root.get("plan_sha256") or ""
    stored_digest = root.get("recovery_sha256") or ""
    if not re.fullmatch(r"[0-9a-f]{64}", source_sha256) or not re.fullmatch(r"[0-9a-f]{64}", plan_sha256):
        raise DocxSemanticsV3Error("recovery map source or plan digest is not a 64-hex value")
    if not re.fullmatch(r"[0-9a-f]{64}", stored_digest):
        raise DocxSemanticsV3Error("recovery map content digest is not a 64-hex value")

    pointers: list[ResolvedV4RecoveryPointer] = []
    projection: dict[str, str] = {}
    bibliography: ResolvedV4RecoveryBibliography | None = None
    for child in root:
        namespace, local = _split_child(child)
        if namespace != RESOLVED_V4_RECOVERY_MAP_NAMESPACE:
            raise DocxSemanticsV3Error("recovery map contains a foreign child element")
        if local == "pointers":
            if pointers or child.text is not None or child.tail is not None:
                raise DocxSemanticsV3Error("recovery map pointers container is not canonical")
            for pointer in child:
                if pointer.tag != f"{{{RESOLVED_V4_RECOVERY_MAP_NAMESPACE}}}pointer":
                    raise DocxSemanticsV3Error("recovery map pointer element is invalid")
                expected = ("role", "relative_path", "bytes", "sha256")
                if tuple(pointer.attrib) != expected or pointer.text is not None or pointer.tail is not None:
                    raise DocxSemanticsV3Error("recovery map pointer attribute order is invalid")
                role = pointer.get("role") or ""
                relative = pointer.get("relative_path") or ""
                byte_text = pointer.get("bytes") or ""
                digest = pointer.get("sha256") or ""
                if role not in RECOVERY_POINTER_ROLES:
                    raise DocxSemanticsV3Error("recovery map pointer role is outside the closed set")
                if not re.fullmatch(r"[0-9a-f]{64}", digest):
                    raise DocxSemanticsV3Error("recovery map pointer digest is not a 64-hex value")
                if not re.fullmatch(r"[1-9][0-9]{0,15}", byte_text):
                    raise DocxSemanticsV3Error("recovery map pointer byte count is not canonical")
                pointers.append(
                    ResolvedV4RecoveryPointer(
                        role,
                        relative,
                        int(byte_text),
                        digest,
                    )
                )
        elif local == "projection":
            if projection or child.text is not None or child.tail is not None:
                raise DocxSemanticsV3Error("recovery map projection container is not canonical")
            expected = ("version", "algorithm", "physical_sha256")
            if tuple(child.attrib) != expected or list(child) or child.text is not None or child.tail is not None:
                raise DocxSemanticsV3Error("recovery map projection attribute order is invalid")
            if child.get("version") != "1":
                raise DocxSemanticsV3Error("recovery map projection version is not 1")
            if child.get("algorithm") != RECOVERY_PROJECTION_ALGORITHM:
                raise DocxSemanticsV3Error("recovery map projection algorithm is not frozen")
            digest = child.get("physical_sha256") or ""
            if not re.fullmatch(r"[0-9a-f]{64}", digest):
                raise DocxSemanticsV3Error("recovery map physical digest is not a 64-hex value")
            projection["physical_sha256"] = digest
        elif local == "bibliography":
            if bibliography is not None or child.text is not None or child.tail is not None:
                raise DocxSemanticsV3Error("recovery map bibliography is not canonical")
            expected = ("owner", "placeholder", "media_type")
            if tuple(child.attrib) != expected or list(child) or child.text is not None or child.tail is not None:
                raise DocxSemanticsV3Error("recovery map bibliography attribute order is invalid")
            bibliography = ResolvedV4RecoveryBibliography(
                child.get("owner") or "",
                child.get("placeholder") or "",
                child.get("media_type") or "",
            )
        else:
            raise DocxSemanticsV3Error("recovery map child order is invalid")

    if not pointers or len(pointers) != len(RECOVERY_POINTER_ROLES):
        raise DocxSemanticsV3Error("recovery map must carry exactly three pointers")
    roles = {item.role for item in pointers}
    if roles != RECOVERY_POINTER_ROLES:
        raise DocxSemanticsV3Error("recovery map pointer roles do not cover the closed set")
    if not projection or bibliography is None:
        raise DocxSemanticsV3Error("recovery map projection or bibliography is missing")
    value = ResolvedV4RecoveryMap(
        source_sha256,
        plan_sha256,
        projection["physical_sha256"],
        tuple(sorted(pointers, key=lambda item: (item.role, item.relative_path))),
        bibliography,
    )
    if _content_digest(value) != stored_digest:
        raise DocxSemanticsV3Error("recovery map content digest does not match its records")
    return value


def _split_child(element: Any) -> tuple[str, str]:
    tag = element.tag
    if not isinstance(tag, str) or tag[0] != "{":
        raise DocxSemanticsV3Error("recovery map child element has no namespace")
    namespace, _, local = tag[1:].partition("}")
    return namespace, local


def _safe_relative(value: str, *, label: str) -> str:
    if not value:
        raise DocxSemanticsV3Error(f"recovery map {label} is empty")
    if value.startswith(("/", "\\")) or re.match(r"^[A-Za-z]:[\\/]", value) or ".." in value.split("/"):
        raise DocxSemanticsV3Error(f"recovery map {label} is not a safe relative path")
    if "\\" in value:
        raise DocxSemanticsV3Error(f"recovery map {label} must use forward slashes")
    if any(part in {".", "..", ""} for part in value.split("/")):
        raise DocxSemanticsV3Error(f"recovery map {label} contains an empty or dot segment")
    return value


def recovery_map_pointers_xml(pointers: tuple[ResolvedV4RecoveryPointer, ...]) -> str:
    return "".join(
        (
            f'<pointer role="{xml_attr(item.role)}" relative_path="{xml_attr(item.relative_path)}" '
            f'bytes="{item.bytes}" sha256="{item.sha256}"/>'
        )
        for item in sorted(pointers, key=lambda entry: (entry.role, entry.relative_path))
    )


def compute_physical_projection(
    path: Path,
    *,
    exclude_item_numbers: set[int],
) -> str:
    """Compute the whole-package physical projection digest.

    The recovery map trio parts, its document relationship, and its content-type
    Overrides are excluded by item number; every other part contributes raw bytes
    and canonical XML.
    """

    with ZipFile(path) as package:
        infos = package.infolist()
        if len({info.filename for info in infos}) != len(infos):
            raise DocxSemanticsV3Error("DOCX package contains duplicate ZIP members")
        members = {info.filename: package.read(info.filename) for info in infos}
    stream: list[bytes] = []
    rels_name = "word/_rels/document.xml.rels"
    types_name = "[Content_Types].xml"
    excluded_parts: set[str] = set()
    for item_number in exclude_item_numbers:
        excluded_parts.add(f"customXml/item{item_number}.xml")
        excluded_parts.add(f"customXml/itemProps{item_number}.xml")
        excluded_parts.add(f"customXml/_rels/item{item_number}.xml.rels")
    for name in sorted(members, key=lambda value: value.encode("utf-8")):
        if name in excluded_parts:
            continue
        if name.endswith(".rels") or name == types_name:
            continue
        raw = members[name]
        stream.extend(_part_stream(name, raw))
    if rels_name in members:
        stream.extend(_filtered_relationship_stream(members[rels_name], exclude_item_numbers))
    if types_name in members:
        stream.extend(_filtered_content_types_stream(members[types_name], exclude_item_numbers))
    return hashlib.sha256(b"".join(stream)).hexdigest()


def _part_stream(name: str, raw: bytes) -> list[bytes]:
    stream = [
        name.encode("utf-8"),
        b"\0",
        str(len(raw)).encode("ascii"),
        b"\0",
        hashlib.sha256(raw).hexdigest().encode("ascii"),
        b"\0",
    ]
    if name.endswith((".xml", ".rels")):
        parser = etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)
        root = etree.fromstring(raw, parser)
        canonical = etree.tostring(
            root,
            method="c14n",
            exclusive=True,
            with_comments=False,
        )
        stream.extend(
            [
                name.encode("utf-8"),
                b"\0",
                str(len(canonical)).encode("ascii"),
                b"\0",
                hashlib.sha256(canonical).hexdigest().encode("ascii"),
                b"\0",
            ]
        )
    return stream


def _filtered_relationship_stream(raw: bytes, exclude_item_numbers: set[int]) -> list[bytes]:
    parser = etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)
    root = etree.fromstring(raw, parser)
    if root.tag != f"{{{_DOC_REL_NS}}}Relationships":
        raise DocxSemanticsV3Error("document relationships root is invalid")
    excluded_targets = {f"../customXml/item{n}.xml" for n in exclude_item_numbers}
    kept = [
        element
        for element in root
        if element.tag == f"{{{_DOC_REL_NS}}}Relationship" and element.get("Target") not in excluded_targets
    ]
    for element in kept:
        if tuple(element.attrib) not in {
            ("Id", "Type", "Target"),
            ("Id", "Type", "Target", "TargetMode"),
        }:
            raise DocxSemanticsV3Error("document relationship attribute order is invalid")
    filtered = etree.Element(f"{{{_DOC_REL_NS}}}Relationships")
    for element in kept:
        filtered.append(element)
    canonical = etree.tostring(filtered, method="c14n", exclusive=True, with_comments=False)
    return [
        b"word/_rels/document.xml.rels",
        b"\0",
        str(len(canonical)).encode("ascii"),
        b"\0",
        hashlib.sha256(canonical).hexdigest().encode("ascii"),
        b"\0",
    ]


def _filtered_content_types_stream(raw: bytes, exclude_item_numbers: set[int]) -> list[bytes]:
    parser = etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)
    root = etree.fromstring(raw, parser)
    if root.tag != f"{{{_CONTENT_TYPES_NS}}}Types":
        raise DocxSemanticsV3Error("content-types root is invalid")
    excluded = {f"/customXml/item{n}.xml" for n in exclude_item_numbers}
    excluded.update({f"/customXml/itemProps{n}.xml" for n in exclude_item_numbers})
    kept = [element for element in root if element.get("PartName") not in excluded]
    filtered = etree.Element(f"{{{_CONTENT_TYPES_NS}}}Types")
    for element in kept:
        filtered.append(element)
    canonical = etree.tostring(filtered, method="c14n", exclusive=True, with_comments=False)
    return [
        b"[Content_Types].xml",
        b"\0",
        str(len(canonical)).encode("ascii"),
        b"\0",
        hashlib.sha256(canonical).hexdigest().encode("ascii"),
        b"\0",
    ]


def inject_recovery_map(path: Path, value: ResolvedV4RecoveryMap) -> None:
    inject_custom_xml_parts(
        path,
        [
            (
                RESOLVED_V4_RECOVERY_MAP_NAMESPACE,
                recovery_map_xml(value),
            )
        ],
        allow_existing_owned=True,
    )


_ITEM_RE = re.compile(r"customXml/item(?P<number>[1-9][0-9]*)\.xml$")


def read_recovery_map(package: ZipFile) -> tuple[int, ResolvedV4RecoveryMap] | None:
    """Return (item_number, map) for the single authenticated recovery map, if present."""

    parser = etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)
    found: tuple[int, ResolvedV4RecoveryMap] | None = None
    for name in package.namelist():
        match = _ITEM_RE.fullmatch(name)
        if match is None:
            continue
        data = package.read(name)
        try:
            root = etree.fromstring(data, parser)
        except etree.XMLSyntaxError:
            continue
        if etree.QName(root).namespace != RESOLVED_V4_RECOVERY_MAP_NAMESPACE:
            continue
        if found is not None:
            raise DocxSemanticsV3Error("duplicate recovery map item part")
        found = (int(match.group("number")), parse_recovery_map(root))
    return found


def build_recovery_map(
    port: Any,
    *,
    input_bytes: ResolvedV4RecoveryInput,
    physical_sha256: str,
) -> ResolvedV4RecoveryMap:
    """Build a frozen recovery map from the authenticated port and raw input bytes."""

    if not re.fullmatch(r"[0-9a-f]{64}", physical_sha256):
        raise DocxSemanticsV3Error("recovery physical digest is not a 64-hex value")
    neutral_name = _safe_relative(input_bytes.neutral_name, label="neutral raw name")
    plan_name = _safe_relative(input_bytes.plan_name, label="plan raw name")
    authored_name = _safe_relative(input_bytes.authored_name, label="authored source name")
    pointers = (
        ResolvedV4RecoveryPointer(
            "neutral_raw",
            neutral_name,
            len(input_bytes.neutral_raw),
            hashlib.sha256(input_bytes.neutral_raw).hexdigest(),
        ),
        ResolvedV4RecoveryPointer(
            "plan_raw",
            plan_name,
            len(input_bytes.plan_raw),
            hashlib.sha256(input_bytes.plan_raw).hexdigest(),
        ),
        ResolvedV4RecoveryPointer(
            "authored_source",
            authored_name,
            len(input_bytes.authored_source),
            hashlib.sha256(input_bytes.authored_source).hexdigest(),
        ),
    )
    source_sha256 = str(getattr(port, "source_sha256", ""))
    plan_sha256 = str(getattr(port, "plan_sha256", ""))
    bibliography = ResolvedV4RecoveryBibliography(
        input_bytes.bibliography_owner,
        input_bytes.bibliography_placeholder,
        input_bytes.bibliography_media_type,
    )
    return ResolvedV4RecoveryMap(
        source_sha256,
        plan_sha256,
        physical_sha256,
        tuple(sorted(pointers, key=lambda item: (item.role, item.relative_path))),
        bibliography,
    )
