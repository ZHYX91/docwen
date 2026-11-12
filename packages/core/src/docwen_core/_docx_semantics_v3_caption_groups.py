"""Persisted ownership of generated ID-less caption/object pairs."""

from __future__ import annotations

import hashlib
from dataclasses import dataclass
from typing import Any

from lxml import etree

from docwen_core._docx_semantics_v3_model import DocxSemanticsV3Error, require_sha256
from docwen_core.docx_caption_sequences import caption_sequence_identifier

PAIRED_CAPTION_MAP_NAMESPACE = "https://docwen.dev/schema/document-paired-caption-map/v2"
PAIRED_CAPTION_TAG_PREFIX = "docwen-caption-pair-v1:"
_KINDS = {"figure", "table", "equation", "code_block"}


@dataclass(frozen=True, slots=True)
class PairedCaptionGroupV3:
    ordinal: int
    kind: str
    tag: str


@dataclass(frozen=True, slots=True)
class PairedCaptionMapV3:
    source_sha256: str
    caption_count: int
    groups: tuple[PairedCaptionGroupV3, ...]
    sequence_scope: str


def derive_paired_caption_group(source_sha256: str, ordinal: int, kind: str) -> PairedCaptionGroupV3:
    require_sha256(source_sha256)
    if type(ordinal) is not int or ordinal < 1 or kind not in _KINDS:
        raise DocxSemanticsV3Error("caption-pair identity has an invalid ordinal or kind")
    digest = hashlib.sha256(f"docwen-caption-pair-v1\0{source_sha256}\0{ordinal}\0{kind}".encode()).hexdigest()
    return PairedCaptionGroupV3(ordinal, kind, f"{PAIRED_CAPTION_TAG_PREFIX}{digest}")


def _validate(record: PairedCaptionMapV3) -> None:
    require_sha256(record.source_sha256)
    try:
        caption_sequence_identifier(record.sequence_scope, "table")
    except ValueError as exc:
        raise DocxSemanticsV3Error("caption-pair map has an invalid sequence scope") from exc
    if type(record.caption_count) is not int or record.caption_count < 1:
        raise DocxSemanticsV3Error("caption-pair map must declare a positive total caption count")
    last = 0
    for group in record.groups:
        if not last < group.ordinal <= record.caption_count:
            raise DocxSemanticsV3Error(
                "caption-pair ordinals are duplicated, unordered or outside the caption inventory"
            )
        if group != derive_paired_caption_group(record.source_sha256, group.ordinal, group.kind):
            raise DocxSemanticsV3Error("caption-pair identity is not canonically derived")
        last = group.ordinal


def paired_caption_map_xml(record: PairedCaptionMapV3) -> bytes:
    _validate(record)
    namespace = PAIRED_CAPTION_MAP_NAMESPACE
    root = etree.Element(
        f"{{{namespace}}}documentPairedCaptionMap",
        nsmap={"dw": namespace},
        version="2",
        source_sha256=record.source_sha256,
        caption_count=str(record.caption_count),
        sequence_scope=record.sequence_scope,
    )
    for group in record.groups:
        etree.SubElement(root, f"{{{namespace}}}pair", ordinal=str(group.ordinal), kind=group.kind, tag=group.tag)
    return etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)


def parse_paired_caption_map(root: Any) -> PairedCaptionMapV3:
    namespace = f"{{{PAIRED_CAPTION_MAP_NAMESPACE}}}"
    count = root.get("caption_count", "")
    if (
        root.tag != f"{namespace}documentPairedCaptionMap"
        or set(root.attrib) != {"version", "source_sha256", "caption_count", "sequence_scope"}
        or root.get("version") != "2"
        or not count.isascii()
        or not count.isdecimal()
        or str(int(count)) != count
        or root.text is not None
        or root.tail is not None
    ):
        raise DocxSemanticsV3Error("caption-pair map root is not closed and canonical")
    groups = []
    for child in root:
        ordinal = child.get("ordinal", "")
        if (
            child.tag != f"{namespace}pair"
            or set(child.attrib) != {"ordinal", "kind", "tag"}
            or not ordinal.isascii()
            or not ordinal.isdecimal()
            or str(int(ordinal)) != ordinal
            or child.text is not None
            or child.tail is not None
            or len(child)
        ):
            raise DocxSemanticsV3Error("caption-pair record is not closed and canonical")
        groups.append(PairedCaptionGroupV3(int(ordinal), child.get("kind"), child.get("tag")))
    record = PairedCaptionMapV3(root.get("source_sha256"), int(count), tuple(groups), root.get("sequence_scope"))
    _validate(record)
    return record
