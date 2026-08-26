"""Frozen identities and value objects for Markdown-semantics v3 DOCX projection."""

from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass
from typing import Any, Literal

TARGET_MAP_NAMESPACE = "https://docwen.dev/schema/document-target-map/v1"
ANCHOR_TOPOLOGY_MAP_NAMESPACE = "https://docwen.dev/schema/document-anchor-topology-map/v1"
SOFT_REFERENCE_MAP_NAMESPACE = "https://docwen.dev/schema/document-soft-reference-map/v1"
REFERENCE_OCCURRENCE_MAP_NAMESPACE = "https://docwen.dev/schema/document-reference-occurrence-map/v1"
CAPTION_STYLE_BINDING_MAP_NAMESPACE = "https://docwen.dev/schema/document-caption-style-binding-map/v1"
TARGET_TAG_PREFIX = "docwen-target-v1:"
ANCHOR_TAG_PREFIX = "docwen-anchor-v1:"
SOFT_REFERENCE_TAG_PREFIX = "docwen-soft-ref-v1:"
REFERENCE_OCCURRENCE_TAG_PREFIX = "docwen-ref-occurrence-v1:"

type TargetKindV3 = Literal["heading", "figure", "table", "equation", "code_block"]
type CaptionStyleKeyV3 = Literal[
    "figure_caption",
    "table_caption",
    "equation_caption",
    "code_block_caption",
]


class DocxSemanticsV3Error(ValueError):
    """Raised when v3 semantic OOXML cannot be emitted or authenticated."""


@dataclass(frozen=True, slots=True)
class TargetIdentityV3:
    kind: TargetKindV3
    source_id: str
    sha256: str
    bookmark_name: str
    tag: str


@dataclass(frozen=True, slots=True)
class AnchorIdentityV3:
    block_kind: str
    source_id: str
    sha256: str
    tag: str


@dataclass(frozen=True, slots=True)
class AnchorTopologyEdgeV3:
    child_tag: str
    parent_tag: str
    sha256: str


@dataclass(frozen=True, slots=True)
class SoftReferenceIdentityV3:
    tag: str
    source_sha256: str
    source_start: int
    source_end: int
    authored_token: str
    cached_number: str


@dataclass(frozen=True, slots=True)
class ReferenceOccurrenceIdentityV3:
    tag: str
    source_sha256: str
    source_start: int
    source_end: int
    authored_token: str
    resolved_bookmark_name: str
    cached_number: str


@dataclass(frozen=True, slots=True)
class CaptionStyleBindingV3:
    """Persisted request-local identity of one managed caption style."""

    semantic_key: CaptionStyleKeyV3
    resolved_style_id: str
    visible_name: str


@dataclass(frozen=True, slots=True)
class SourceAnchorV3:
    owner_kind: Literal["semantic_target", "ordinary_anchor"]
    source_id: str
    block_kind: str


@dataclass(frozen=True, slots=True)
class OrdinaryAnchorGroupV3:
    """Authenticated logical group owned by one ordinary Markdown anchor."""

    anchor: SourceAnchorV3
    elements: tuple[Any, ...]
    index: int


@dataclass(frozen=True, slots=True)
class TargetBindingV3:
    identity: TargetIdentityV3
    elements: tuple[Any, ...]


@dataclass(frozen=True, slots=True)
class AnchorBindingV3:
    identity: AnchorIdentityV3
    elements: tuple[Any, ...]
    direct_parent_source_id: str | None = None


@dataclass(frozen=True, slots=True)
class CaptionBindingV3:
    kind: Literal["figure", "table", "equation", "code_block"]
    caption_element: Any
    object_elements: tuple[Any, ...]
    source_id: str | None
    title: str
    cached_number: str


@dataclass(frozen=True, slots=True)
class RecoveredCaptionV3:
    kind: Literal["figure", "table", "equation", "code_block"]
    source_id: str | None
    title: str
    cached_number: str
    caption_element: Any
    object_elements: tuple[Any, ...]


def derive_target_identity_v3(kind: TargetKindV3, source_id: str) -> TargetIdentityV3:
    """Derive the frozen complete target hash, bookmark, and outer SDT tag."""

    if kind not in {"heading", "figure", "table", "equation", "code_block"}:
        raise DocxSemanticsV3Error(f"unsupported v3 target kind: {kind}")
    require_source_id(source_id)
    digest = hashlib.sha256(f"docwen-target-map-v1\0{kind}\0{source_id}".encode()).hexdigest()
    return TargetIdentityV3(
        kind=kind,
        source_id=source_id,
        sha256=digest,
        bookmark_name=f"DW_T_{digest[:35]}",
        tag=f"{TARGET_TAG_PREFIX}{digest[:32]}",
    )


def derive_anchor_identity_v3(block_kind: str, source_id: str) -> AnchorIdentityV3:
    """Derive an opaque ordinary-anchor tag; block kind is metadata only."""

    if block_kind not in {
        "paragraph",
        "image",
        "table",
        "equation",
        "code_block",
        "list",
        "list_item",
        "block_quote",
        "callout",
    }:
        raise DocxSemanticsV3Error(f"unsupported v3 ordinary-anchor block kind: {block_kind}")
    require_source_id(source_id)
    digest = hashlib.sha256(f"docwen-anchor-map-v1\0anchor\0{source_id}".encode()).hexdigest()
    return AnchorIdentityV3(
        block_kind=block_kind,
        source_id=source_id,
        sha256=digest,
        tag=f"{ANCHOR_TAG_PREFIX}{digest[:32]}",
    )


def derive_anchor_topology_edge_v3(child_tag: str, parent_tag: str) -> AnchorTopologyEdgeV3:
    """Derive one closed direct ordinary-anchor topology edge."""

    pattern = rf"{re.escape(ANCHOR_TAG_PREFIX)}[0-9a-f]{{32}}"
    if re.fullmatch(pattern, child_tag) is None or re.fullmatch(pattern, parent_tag) is None:
        raise DocxSemanticsV3Error("anchor-topology endpoints must be canonical ordinary-anchor tags")
    if child_tag == parent_tag:
        raise DocxSemanticsV3Error("anchor-topology edge must not be self-referential")
    digest = hashlib.sha256(f"docwen-anchor-topology-edge-v1\0{child_tag}\0{parent_tag}".encode()).hexdigest()
    return AnchorTopologyEdgeV3(child_tag=child_tag, parent_tag=parent_tag, sha256=digest)


def derive_soft_reference_identity_v3(
    *,
    source_sha256: str,
    source_start: int,
    source_end: int,
    authored_token: str,
    cached_number: str,
) -> SoftReferenceIdentityV3:
    """Derive the frozen inline non-target SDT identity."""

    require_sha256(source_sha256)
    if source_start < 0 or source_end <= source_start:
        raise DocxSemanticsV3Error("soft-reference source range must be non-empty and ordered")
    if not authored_token.startswith("@[[") or not authored_token.endswith("]]"):
        raise DocxSemanticsV3Error("soft-reference authored token is not canonical")
    if not cached_number:
        raise DocxSemanticsV3Error("soft-reference cached number must be non-empty")
    if source_end - source_start != len(authored_token):
        raise DocxSemanticsV3Error("soft-reference range length does not match its authored token")
    preimage = f"docwen-soft-ref-map-v1\0{source_sha256}\0{source_start}\0{source_end}\0{authored_token}"
    digest = hashlib.sha256(preimage.encode("utf-8")).hexdigest()
    return SoftReferenceIdentityV3(
        tag=f"{SOFT_REFERENCE_TAG_PREFIX}{digest[:32]}",
        source_sha256=source_sha256,
        source_start=source_start,
        source_end=source_end,
        authored_token=authored_token,
        cached_number=cached_number,
    )


def derive_reference_occurrence_identity_v3(
    *,
    source_sha256: str,
    source_start: int,
    source_end: int,
    authored_token: str,
    resolved_bookmark_name: str,
    cached_number: str,
) -> ReferenceOccurrenceIdentityV3:
    """Derive one authenticated REF-based authored occurrence identity."""

    require_sha256(source_sha256)
    if source_start < 0 or source_end <= source_start:
        raise DocxSemanticsV3Error("reference-occurrence source range must be non-empty and ordered")
    if not authored_token.startswith("@[[") or not authored_token.endswith("]]"):
        raise DocxSemanticsV3Error("reference-occurrence authored token is not canonical")
    if source_end - source_start != len(authored_token):
        raise DocxSemanticsV3Error("reference-occurrence range length does not match its authored token")
    if re.fullmatch(r"DW_T_[0-9a-f]{35}", resolved_bookmark_name) is None:
        raise DocxSemanticsV3Error("reference-occurrence bookmark name is not canonical")
    if not cached_number:
        raise DocxSemanticsV3Error("reference-occurrence cached number must be non-empty")
    preimage = (
        f"docwen-ref-occurrence-map-v1\0{source_sha256}\0{source_start}\0{source_end}\0"
        f"{authored_token}\0{resolved_bookmark_name}\0{cached_number}"
    )
    digest = hashlib.sha256(preimage.encode("utf-8")).hexdigest()
    return ReferenceOccurrenceIdentityV3(
        tag=f"{REFERENCE_OCCURRENCE_TAG_PREFIX}{digest[:32]}",
        source_sha256=source_sha256,
        source_start=source_start,
        source_end=source_end,
        authored_token=authored_token,
        resolved_bookmark_name=resolved_bookmark_name,
        cached_number=cached_number,
    )


def require_source_id(source_id: str) -> None:
    if re.fullmatch(r"[A-Za-z0-9-]{1,128}", source_id) is None:
        raise DocxSemanticsV3Error("source ID must match [A-Za-z0-9-]{1,128}")


def require_sha256(value: str) -> None:
    if re.fullmatch(r"[0-9a-f]{64}", value) is None:
        raise DocxSemanticsV3Error("SHA-256 must be lowercase 64-hex")
