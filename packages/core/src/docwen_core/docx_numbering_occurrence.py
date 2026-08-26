"""Closed source-recovery authority for disabled ID-less v4 captions."""

from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass
from itertools import pairwise
from typing import Any, Literal

from docwen_core._docx_semantics_v3_model import DocxSemanticsV3Error, require_sha256
from docwen_core._docx_semantics_v3_ooxml import sdt_tag, wrap_direct_body_group

NUMBERING_OCCURRENCE_MAP_NAMESPACE = "https://docwen.dev/schema/document-numbering-occurrence-map/v1"
NUMBERING_OCCURRENCE_TAG_PREFIX = "docwen-numbering-occurrence-v1:"
_XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

type CaptionKind = Literal["figure", "table", "equation", "code_block"]


@dataclass(frozen=True, slots=True)
class NumberingOccurrenceIdentity:
    """One authenticated disabled ID-less declaration/object occurrence."""

    tag: str
    source_sha256: str
    source_start: int
    source_end: int
    kind: CaptionKind
    plan_sha256: str
    sha256: str
    enabled: Literal[False] = False
    target_id: None = None
    derived_number: None = None


def derive_numbering_occurrence(
    *,
    source_sha256: str,
    source_start: int,
    source_end: int,
    kind: CaptionKind,
    plan_sha256: str,
) -> NumberingOccurrenceIdentity:
    """Derive the spec-frozen full digest and physical SDT tag."""

    require_sha256(source_sha256)
    require_sha256(plan_sha256)
    if source_start < 0 or source_end <= source_start:
        raise DocxSemanticsV3Error("numbering-occurrence source range must be non-empty and ordered")
    if kind not in {"figure", "table", "equation", "code_block"}:
        raise DocxSemanticsV3Error("numbering-occurrence kind is invalid")
    preimage = (
        "docwen-numbering-occurrence-map-v1\0"
        f"{source_sha256}\0{source_start}\0{source_end}\0{kind}\0false\0\0\0{plan_sha256}"
    )
    digest = hashlib.sha256(preimage.encode("utf-8")).hexdigest()
    return NumberingOccurrenceIdentity(
        tag=f"{NUMBERING_OCCURRENCE_TAG_PREFIX}{digest[:32]}",
        source_sha256=source_sha256,
        source_start=source_start,
        source_end=source_end,
        kind=kind,
        plan_sha256=plan_sha256,
        sha256=digest,
    )


def numbering_occurrence_map_xml(records: list[NumberingOccurrenceIdentity]) -> bytes:
    """Serialize one non-empty canonical map in authenticated source order."""

    validated = validate_numbering_occurrences(records)
    plan_sha256 = validated[0].plan_sha256
    entries = "".join(
        (
            f'<occurrence tag="{item.tag}" source_sha256="{item.source_sha256}" '
            f'source_start="{item.source_start}" source_end="{item.source_end}" '
            f'kind="{item.kind}" enabled="false" target_id="" derived_number="" '
            f'plan_sha256="{item.plan_sha256}" sha256="{item.sha256}"/>'
        )
        for item in validated
    )
    root = (
        f'<documentNumberingOccurrenceMap xmlns="{NUMBERING_OCCURRENCE_MAP_NAMESPACE}" '
        f'version="1" plan_sha256="{plan_sha256}">{entries}'
        "</documentNumberingOccurrenceMap>"
    )
    return f"{_XML_DECLARATION}\n{root}\n".encode()


def parse_numbering_occurrence_map(root: Any) -> list[NumberingOccurrenceIdentity]:
    """Parse and recompute every record; lexical bytes are checked by caller."""

    namespace = f"{{{NUMBERING_OCCURRENCE_MAP_NAMESPACE}}}"
    if (
        root.tag != f"{namespace}documentNumberingOccurrenceMap"
        or tuple(root.attrib) != ("version", "plan_sha256")
        or root.get("version") != "1"
        or root.text is not None
        or root.tail is not None
    ):
        raise DocxSemanticsV3Error("numbering-occurrence root is not canonical")
    plan_sha256 = root.get("plan_sha256", "")
    require_sha256(plan_sha256)
    records: list[NumberingOccurrenceIdentity] = []
    expected_attributes = (
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
    for element in root:
        if (
            element.tag != f"{namespace}occurrence"
            or tuple(element.attrib) != expected_attributes
            or element.text is not None
            or element.tail is not None
            or len(element) != 0
            or element.get("enabled") != "false"
            or element.get("target_id") != ""
            or element.get("derived_number") != ""
            or element.get("plan_sha256") != plan_sha256
        ):
            raise DocxSemanticsV3Error("numbering-occurrence record is not canonical")
        try:
            source_start = int(element.get("source_start", ""))
            source_end = int(element.get("source_end", ""))
        except ValueError as exc:
            raise DocxSemanticsV3Error("numbering-occurrence range is invalid") from exc
        kind = element.get("kind", "")
        if kind not in {"figure", "table", "equation", "code_block"}:
            raise DocxSemanticsV3Error("numbering-occurrence kind is invalid")
        derived = derive_numbering_occurrence(
            source_sha256=element.get("source_sha256", ""),
            source_start=source_start,
            source_end=source_end,
            kind=kind,  # type: ignore[arg-type]
            plan_sha256=plan_sha256,
        )
        if element.get("tag") != derived.tag or element.get("sha256") != derived.sha256:
            raise DocxSemanticsV3Error("numbering-occurrence digest is invalid")
        records.append(derived)
    return list(validate_numbering_occurrences(records))


def validate_numbering_occurrences(
    records: list[NumberingOccurrenceIdentity],
) -> tuple[NumberingOccurrenceIdentity, ...]:
    if not records:
        raise DocxSemanticsV3Error("numbering-occurrence map must not be empty")
    canonical = sorted(records, key=lambda item: (item.source_start, item.source_end, item.kind, item.tag))
    if records != canonical:
        raise DocxSemanticsV3Error("numbering-occurrence records are not canonically ordered")
    if len({item.tag for item in records}) != len(records):
        raise DocxSemanticsV3Error("numbering-occurrence tags are not unique")
    if len({item.plan_sha256 for item in records}) != 1:
        raise DocxSemanticsV3Error("numbering-occurrence records mix plan identities")
    for item in records:
        expected = derive_numbering_occurrence(
            source_sha256=item.source_sha256,
            source_start=item.source_start,
            source_end=item.source_end,
            kind=item.kind,
            plan_sha256=item.plan_sha256,
        )
        if item != expected:
            raise DocxSemanticsV3Error("numbering-occurrence identity is not canonically derived")
    for previous, current in pairwise(records):
        if current.source_start < previous.source_end:
            raise DocxSemanticsV3Error("numbering-occurrence source ranges overlap")
    return tuple(records)


def wrap_numbering_occurrence(
    caption_element: Any,
    object_elements: tuple[Any, ...],
    identity: NumberingOccurrenceIdentity,
) -> None:
    """Wrap exactly two logical slots in the kind-specific physical order."""

    if not object_elements:
        raise DocxSemanticsV3Error("numbering occurrence has no logical object")
    object_owner = _one_direct_body_owner(object_elements)
    caption_owner = _direct_body_owner(caption_element)
    if object_owner is caption_owner:
        raise DocxSemanticsV3Error("caption and object must occupy distinct physical slots")
    elements = (object_owner, caption_owner) if identity.kind == "figure" else (caption_owner, object_owner)
    wrap_direct_body_group(elements, identity.tag)


def prove_numbering_occurrence_sdt(
    sdt: Any,
    identity: NumberingOccurrenceIdentity,
    *,
    caption_style_id: str,
    allowed_inline_tags: tuple[str, ...] = (),
) -> tuple[Any, Any]:
    """Prove the exact two-slot, field-free physical occurrence envelope."""

    from docx.oxml.ns import qn

    if sdt.tag != qn("w:sdt") or sdt_tag(sdt) != identity.tag:
        raise DocxSemanticsV3Error("numbering-occurrence SDT tag is invalid")
    properties = sdt.find(qn("w:sdtPr"))
    content = sdt.find(qn("w:sdtContent"))
    if (
        properties is None
        or content is None
        or [item.tag for item in list(sdt)] != [qn("w:sdtPr"), qn("w:sdtContent")]
        or len(content) != 2
    ):
        raise DocxSemanticsV3Error("numbering-occurrence SDT envelope is not canonical")
    blocks = tuple(content)
    caption_index = 1 if identity.kind == "figure" else 0
    caption = blocks[caption_index]
    logical_object = blocks[1 - caption_index]
    if caption.tag != qn("w:p"):
        raise DocxSemanticsV3Error("numbering-occurrence caption slot is not a paragraph")
    styles = caption.findall(f"{qn('w:pPr')}/{qn('w:pStyle')}")
    if len(styles) != 1 or styles[0].get(qn("w:val")) != caption_style_id:
        raise DocxSemanticsV3Error("numbering-occurrence caption style is not exact")
    inline_carriers = list(caption.iter(qn("w:sdt")))
    inline_tags = [sdt_tag(item) for item in inline_carriers]
    if inline_tags != list(allowed_inline_tags) or any(item.getparent() is not caption for item in inline_carriers):
        raise DocxSemanticsV3Error("numbering-occurrence inline carrier order is not exact")
    allowed = set(inline_carriers)
    bookmarks = [
        item
        for name in ("w:bookmarkStart", "w:bookmarkEnd")
        for item in sdt.iter(qn(name))
        if not _is_inside_allowed_carrier(item, allowed)
    ]
    if bookmarks:
        raise DocxSemanticsV3Error("disabled ID-less occurrence must not contain a bookmark")
    instructions = "".join(
        item.text or "" for item in sdt.iter(qn("w:instrText")) if not _is_inside_allowed_carrier(item, allowed)
    )
    if re.search(r"\b(?:SEQ|REF|CITATION)\b", instructions, re.IGNORECASE):
        raise DocxSemanticsV3Error("disabled ID-less occurrence contains an unowned field")
    _prove_object_slot(logical_object, identity.kind)
    return caption, logical_object


def _is_inside_allowed_carrier(element: Any, allowed: set[Any]) -> bool:
    current = element.getparent()
    while current is not None:
        if current in allowed:
            return True
        current = current.getparent()
    return False


def _prove_object_slot(element: Any, kind: CaptionKind) -> None:
    from docx.oxml.ns import qn

    logical = _unwrap_ordinary_anchor(element)
    if kind == "table" and logical.tag != qn("w:tbl"):
        raise DocxSemanticsV3Error("table occurrence has no table object")
    if kind == "figure" and not list(logical.iter(qn("w:drawing"))):
        raise DocxSemanticsV3Error("figure occurrence has no drawing object")
    if kind == "equation" and not (
        list(logical.iter("{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath"))
        or list(logical.iter("{http://schemas.openxmlformats.org/officeDocument/2006/math}oMathPara"))
    ):
        raise DocxSemanticsV3Error("equation occurrence has no OMML object")
    if kind == "code_block" and logical.tag not in {qn("w:p"), qn("w:sdt")}:
        raise DocxSemanticsV3Error("code-block occurrence has no logical block object")


def _unwrap_ordinary_anchor(element: Any) -> Any:
    from docx.oxml.ns import qn

    if element.tag != qn("w:sdt"):
        return element
    tag = sdt_tag(element) or ""
    if not tag.startswith("docwen-anchor-v1:"):
        return element
    content = element.find(qn("w:sdtContent"))
    if content is None or len(content) != 1:
        raise DocxSemanticsV3Error("ordinary-anchor object slot is not singular")
    return content[0]


def _one_direct_body_owner(elements: tuple[Any, ...]) -> Any:
    owners = tuple(dict.fromkeys(_direct_body_owner(element) for element in elements))
    if len(owners) != 1:
        raise DocxSemanticsV3Error("numbering occurrence object must occupy exactly one logical slot")
    return owners[0]


def _direct_body_owner(element: Any) -> Any:
    from docx.oxml.ns import qn

    owner = element
    while owner.getparent() is not None and owner.getparent().tag != qn("w:body"):
        owner = owner.getparent()
    if owner.getparent() is None or owner.getparent().tag != qn("w:body"):
        raise DocxSemanticsV3Error("numbering occurrence is detached from the document body")
    return owner


__all__ = [
    "NUMBERING_OCCURRENCE_MAP_NAMESPACE",
    "NUMBERING_OCCURRENCE_TAG_PREFIX",
    "NumberingOccurrenceIdentity",
    "derive_numbering_occurrence",
    "numbering_occurrence_map_xml",
    "parse_numbering_occurrence_map",
    "prove_numbering_occurrence_sdt",
    "validate_numbering_occurrences",
    "wrap_numbering_occurrence",
]
