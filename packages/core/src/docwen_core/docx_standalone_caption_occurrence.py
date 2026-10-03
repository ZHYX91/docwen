"""Closed resolved-v4 authority for ID-less standalone captions."""

from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass
from itertools import pairwise
from typing import Any, Literal

from docwen_core._docx_semantics_v3_model import DocxSemanticsV3Error, require_sha256
from docwen_core._docx_semantics_v3_ooxml import sdt_tag, wrap_direct_body_group

STANDALONE_CAPTION_OCCURRENCE_MAP_NAMESPACE = "https://docwen.dev/schema/document-standalone-caption-occurrence-map/v1"
STANDALONE_CAPTION_OCCURRENCE_TAG_PREFIX = "docwen-standalone-caption-v1:"
_XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

type CaptionKind = Literal["figure", "table", "equation", "code_block"]


@dataclass(frozen=True, slots=True)
class StandaloneCaptionOccurrenceIdentity:
    """One authenticated ID-less standalone caption occurrence."""

    tag: str
    source_sha256: str
    source_start: int
    source_end: int
    kind: CaptionKind
    plan_sha256: str
    enabled: bool
    target_id: None
    derived_number: str | None
    sha256: str


def derive_standalone_caption_occurrence(
    *,
    source_sha256: str,
    source_start: int,
    source_end: int,
    kind: CaptionKind,
    plan_sha256: str,
    enabled: bool,
    derived_number: str | None,
) -> StandaloneCaptionOccurrenceIdentity:
    """Derive the canonical digest and block-SDT tag."""

    require_sha256(source_sha256)
    require_sha256(plan_sha256)
    if source_start < 0 or source_end <= source_start:
        raise DocxSemanticsV3Error("standalone-caption occurrence source range must be non-empty and ordered")
    if kind not in {"figure", "table", "equation", "code_block"}:
        raise DocxSemanticsV3Error("standalone-caption occurrence kind is invalid")
    if type(enabled) is not bool:
        raise DocxSemanticsV3Error("standalone-caption occurrence enabled flag is invalid")
    if enabled:
        if not isinstance(derived_number, str) or not derived_number:
            raise DocxSemanticsV3Error("enabled standalone-caption occurrence requires a derived number")
    elif derived_number is not None:
        raise DocxSemanticsV3Error("disabled standalone-caption occurrence must not carry a derived number")
    encoded_number = derived_number or ""
    encoded_enabled = "true" if enabled else "false"
    preimage = (
        "docwen-standalone-caption-occurrence-map-v1\0"
        f"{source_sha256}\0{source_start}\0{source_end}\0{kind}\0"
        f"{encoded_enabled}\0\0{encoded_number}\0{plan_sha256}"
    )
    digest = hashlib.sha256(preimage.encode("utf-8")).hexdigest()
    return StandaloneCaptionOccurrenceIdentity(
        tag=f"{STANDALONE_CAPTION_OCCURRENCE_TAG_PREFIX}{digest[:32]}",
        source_sha256=source_sha256,
        source_start=source_start,
        source_end=source_end,
        kind=kind,
        plan_sha256=plan_sha256,
        enabled=enabled,
        target_id=None,
        derived_number=derived_number,
        sha256=digest,
    )


def standalone_caption_occurrence_map_xml(
    records: list[StandaloneCaptionOccurrenceIdentity],
) -> bytes:
    """Serialize a non-empty canonical source-order map."""

    validated = validate_standalone_caption_occurrences(records)
    plan_sha256 = validated[0].plan_sha256
    entries = "".join(
        (
            f'<occurrence tag="{item.tag}" source_sha256="{item.source_sha256}" '
            f'source_start="{item.source_start}" source_end="{item.source_end}" '
            f'kind="{item.kind}" enabled="{"true" if item.enabled else "false"}" target_id="" '
            f'derived_number="{_xml_attr(item.derived_number or "")}" '
            f'plan_sha256="{item.plan_sha256}" sha256="{item.sha256}"/>'
        )
        for item in validated
    )
    root = (
        f'<documentStandaloneCaptionOccurrenceMap xmlns="{STANDALONE_CAPTION_OCCURRENCE_MAP_NAMESPACE}" '
        f'version="1" plan_sha256="{plan_sha256}">{entries}'
        "</documentStandaloneCaptionOccurrenceMap>"
    )
    return f"{_XML_DECLARATION}\n{root}\n".encode()


def parse_standalone_caption_occurrence_map(root: Any) -> list[StandaloneCaptionOccurrenceIdentity]:
    """Parse and rederive every canonical map record."""

    namespace = f"{{{STANDALONE_CAPTION_OCCURRENCE_MAP_NAMESPACE}}}"
    if (
        root.tag != f"{namespace}documentStandaloneCaptionOccurrenceMap"
        or set(root.attrib) != {"version", "plan_sha256"}
        or root.get("version") != "1"
        or root.text is not None
        or root.tail is not None
    ):
        raise DocxSemanticsV3Error("standalone-caption occurrence root is not canonical")
    plan_sha256 = root.get("plan_sha256", "")
    require_sha256(plan_sha256)
    expected_attributes = {
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
    }
    records: list[StandaloneCaptionOccurrenceIdentity] = []
    for element in root:
        if (
            element.tag != f"{namespace}occurrence"
            or set(element.attrib) != expected_attributes
            or element.text is not None
            or element.tail is not None
            or len(element) != 0
            or element.get("target_id") != ""
            or element.get("plan_sha256") != plan_sha256
        ):
            raise DocxSemanticsV3Error("standalone-caption occurrence record is not canonical")
        raw_enabled = element.get("enabled")
        if raw_enabled not in {"true", "false"}:
            raise DocxSemanticsV3Error("standalone-caption occurrence enabled flag is invalid")
        enabled = raw_enabled == "true"
        derived_number = element.get("derived_number", "")
        if not derived_number:
            resolved_number: str | None = None
        else:
            resolved_number = derived_number
        try:
            source_start = int(element.get("source_start", ""))
            source_end = int(element.get("source_end", ""))
        except ValueError as exc:
            raise DocxSemanticsV3Error("standalone-caption occurrence range is invalid") from exc
        kind = element.get("kind", "")
        if kind not in {"figure", "table", "equation", "code_block"}:
            raise DocxSemanticsV3Error("standalone-caption occurrence kind is invalid")
        derived = derive_standalone_caption_occurrence(
            source_sha256=element.get("source_sha256", ""),
            source_start=source_start,
            source_end=source_end,
            kind=kind,  # type: ignore[arg-type]
            plan_sha256=plan_sha256,
            enabled=enabled,
            derived_number=resolved_number,
        )
        if element.get("tag") != derived.tag or element.get("sha256") != derived.sha256:
            raise DocxSemanticsV3Error("standalone-caption occurrence digest is invalid")
        records.append(derived)
    return list(validate_standalone_caption_occurrences(records))


def validate_standalone_caption_occurrences(
    records: list[StandaloneCaptionOccurrenceIdentity],
) -> tuple[StandaloneCaptionOccurrenceIdentity, ...]:
    if not records:
        raise DocxSemanticsV3Error("standalone-caption occurrence map must not be empty")
    canonical = sorted(records, key=lambda item: (item.source_start, item.source_end, item.kind, item.tag))
    if records != canonical:
        raise DocxSemanticsV3Error("standalone-caption occurrence records are not canonically ordered")
    if len({item.tag for item in records}) != len(records):
        raise DocxSemanticsV3Error("standalone-caption occurrence tags are not unique")
    if len({item.plan_sha256 for item in records}) != 1:
        raise DocxSemanticsV3Error("standalone-caption occurrence records mix plan identities")
    for item in records:
        expected = derive_standalone_caption_occurrence(
            source_sha256=item.source_sha256,
            source_start=item.source_start,
            source_end=item.source_end,
            kind=item.kind,
            plan_sha256=item.plan_sha256,
            enabled=item.enabled,
            derived_number=item.derived_number,
        )
        if item != expected:
            raise DocxSemanticsV3Error("standalone-caption occurrence identity is not canonically derived")
    for previous, current in pairwise(records):
        if current.source_start < previous.source_end:
            raise DocxSemanticsV3Error("standalone-caption occurrence source ranges overlap")
    return tuple(records)


def wrap_standalone_caption_occurrence(
    caption_element: Any,
    identity: StandaloneCaptionOccurrenceIdentity,
) -> None:
    """Wrap exactly one direct-body caption paragraph."""

    owner = _direct_body_owner(caption_element)
    wrap_direct_body_group((owner,), identity.tag)


def prove_standalone_caption_occurrence_sdt(
    sdt: Any,
    identity: StandaloneCaptionOccurrenceIdentity,
    *,
    caption_style_id: str,
    allowed_inline_tags: tuple[str, ...] = (),
) -> Any:
    """Prove one ID-less caption-only physical occurrence envelope."""

    from docx.oxml.ns import qn

    if sdt.tag != qn("w:sdt") or sdt_tag(sdt) != identity.tag:
        raise DocxSemanticsV3Error("standalone-caption occurrence SDT tag is invalid")
    properties = sdt.find(qn("w:sdtPr"))
    content = sdt.find(qn("w:sdtContent"))
    if (
        properties is None
        or content is None
        or [item.tag for item in list(sdt)] != [qn("w:sdtPr"), qn("w:sdtContent")]
        or len(content) != 1
    ):
        raise DocxSemanticsV3Error("standalone-caption occurrence SDT envelope is not canonical")
    caption = content[0]
    if caption.tag != qn("w:p"):
        raise DocxSemanticsV3Error("standalone-caption occurrence slot is not a paragraph")
    styles = caption.findall(f"{qn('w:pPr')}/{qn('w:pStyle')}")
    if len(styles) != 1 or styles[0].get(qn("w:val")) != caption_style_id:
        raise DocxSemanticsV3Error("standalone-caption occurrence style is not exact")
    inline_carriers = list(caption.iter(qn("w:sdt")))
    inline_tags = [sdt_tag(item) for item in inline_carriers]
    if inline_tags != list(allowed_inline_tags) or any(item.getparent() is not caption for item in inline_carriers):
        raise DocxSemanticsV3Error("standalone-caption occurrence inline carrier order is not exact")
    allowed = set(inline_carriers)
    bookmarks = [
        item
        for name in ("w:bookmarkStart", "w:bookmarkEnd")
        for item in sdt.iter(qn(name))
        if not _is_inside_allowed_carrier(item, allowed)
    ]
    if bookmarks:
        raise DocxSemanticsV3Error("ID-less standalone caption must not contain a bookmark")
    instructions = "".join(
        item.text or "" for item in sdt.iter(qn("w:instrText")) if not _is_inside_allowed_carrier(item, allowed)
    )
    forbidden = r"\b(?:REF|CITATION)\b" if identity.enabled else r"\b(?:SEQ|STYLEREF|REF|CITATION)\b"
    if re.search(forbidden, instructions, re.IGNORECASE):
        raise DocxSemanticsV3Error("standalone-caption occurrence contains unowned fields")
    return caption


def _is_inside_allowed_carrier(element: Any, allowed: set[Any]) -> bool:
    current = element.getparent()
    while current is not None:
        if current in allowed:
            return True
        current = current.getparent()
    return False


def _direct_body_owner(element: Any) -> Any:
    from docx.oxml.ns import qn

    owner = element
    while owner.getparent() is not None and owner.getparent().tag != qn("w:body"):
        owner = owner.getparent()
    if owner.getparent() is None or owner.getparent().tag != qn("w:body"):
        raise DocxSemanticsV3Error("standalone-caption occurrence is detached from the document body")
    return owner


def _xml_attr(value: str) -> str:
    return (
        value.replace("&", "&amp;")
        .replace('"', "&quot;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace("\r", "&#13;")
        .replace("\n", "&#10;")
        .replace("\t", "&#9;")
    )


__all__ = [
    "STANDALONE_CAPTION_OCCURRENCE_MAP_NAMESPACE",
    "STANDALONE_CAPTION_OCCURRENCE_TAG_PREFIX",
    "StandaloneCaptionOccurrenceIdentity",
    "derive_standalone_caption_occurrence",
    "parse_standalone_caption_occurrence_map",
    "prove_standalone_caption_occurrence_sdt",
    "standalone_caption_occurrence_map_xml",
    "validate_standalone_caption_occurrences",
    "wrap_standalone_caption_occurrence",
]
