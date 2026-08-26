"""Closed v4 resolved-document and structured-numbering Conversion Port.

The port is provider-neutral.  It carries authored Markdown plus source-bound
semantic occurrences in one resource and presentation-only numbering facts in
the other.  It never carries a resolver, Workspace object, or a hidden
options bag.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Literal

RESOLVED_DOCUMENT_SCHEMA = "docwen.resolved_document.v1"
RESOLVED_DOCUMENT_SCHEMA_ID = "urn:docwen:schema:resolved-document:v1"
RESOLVED_DOCUMENT_MEDIA_TYPE = "application/vnd.docwen.resolved-document+json"
NUMBERING_EXPORT_PLAN_SCHEMA = "docwen.numbering_export_plan.v1"
NUMBERING_EXPORT_PLAN_SCHEMA_ID = "urn:docwen:schema:numbering-export-plan:v1"
NUMBERING_EXPORT_PLAN_MEDIA_TYPE = "application/vnd.docwen.numbering-export-plan+json"
MAX_RESOLVED_NUMBERING_RESOURCE_BYTES = 8 * 1024 * 1024
MAX_RESOLVED_DOCUMENT_EMBEDDED_BYTES = 6_000_000

type TargetKind = Literal["heading", "figure", "table", "equation", "code_block"]
type HeadingNumberFormat = Literal[
    "chinese_lower",
    "chinese_upper",
    "arabic_half",
    "arabic_full",
    "arabic_circled",
    "letter_upper",
    "letter_lower",
    "roman_upper",
    "roman_lower",
]
type CaptionNumberFormat = Literal["arabic_half", "letter_upper", "letter_lower", "roman_upper", "roman_lower"]


class ResolvedNumberingPortError(ValueError):
    """Stable fail-closed admission error for the dual-input port."""

    def __init__(self, code: str, message: str) -> None:
        self.code = code
        super().__init__(message)


@dataclass(frozen=True, slots=True)
class HeadingCounterSegment:
    level: int
    number_format: HeadingNumberFormat


@dataclass(frozen=True, slots=True)
class HeadingLiteralSegment:
    literal: str


type HeadingDisplaySegment = HeadingCounterSegment | HeadingLiteralSegment


@dataclass(frozen=True, slots=True)
class HeadingLevelDefinition:
    level: int
    start: int
    number_format: HeadingNumberFormat
    display: tuple[HeadingDisplaySegment, ...]
    suffix: Literal["nothing", "space", "tab"]
    restart_after_level: int | None


@dataclass(frozen=True, slots=True)
class HeadingDefinition:
    definition_id: str
    levels: tuple[HeadingLevelDefinition, ...]


@dataclass(frozen=True, slots=True)
class HeadingStart:
    level: int
    value: int


@dataclass(frozen=True, slots=True)
class HeadingInstance:
    instance_id: str
    definition_id: str
    starts: tuple[HeadingStart, ...]


@dataclass(frozen=True, slots=True)
class HeadingListMaterialization:
    definition_id: str
    instance_id: str
    level: int
    type: Literal["heading_list"] = "heading_list"


@dataclass(frozen=True, slots=True)
class CaptionMaterialization:
    type: Literal["simple_seq", "chapter_seq"]
    counter: Literal["Figure", "Table", "Equation", "Code"]
    number_format: CaptionNumberFormat
    sequence_action: Literal["continue", "reset_to_start", "restart_by_heading_level"]
    start_value: int | None
    chapter_heading_level: int | None
    chapter_heading_style: str | None
    chapter_separator: str | None
    restart_heading_level: int | None
    restart_heading_style: str | None
    chapter_cached_number: str | None
    sequence_cached_number: str
    localized_label: str
    label_separator: str


type TargetMaterialization = HeadingListMaterialization | CaptionMaterialization


@dataclass(frozen=True, slots=True)
class ResolvedDocumentTarget:
    source_start: int
    source_end: int
    source_slice_sha256: str
    kind: TargetKind
    target_id: str | None
    heading_level: int | None
    authored_text: str

    @property
    def occurrence_key(self) -> tuple[int, int, TargetKind]:
        return (self.source_start, self.source_end, self.kind)


@dataclass(frozen=True, slots=True)
class ResolvedReference:
    source_start: int
    source_end: int
    source_slice_sha256: str
    authored_token: str
    target_source_start: int
    target_source_end: int
    target_kind: TargetKind
    target_id: str | None
    cached_number: str
    alias: str | None

    @property
    def target_occurrence_key(self) -> tuple[int, int, TargetKind]:
        return (self.target_source_start, self.target_source_end, self.target_kind)


@dataclass(frozen=True, slots=True)
class ResolvedResourceOccurrence:
    source_start: int
    source_end: int
    source_slice_sha256: str
    authored_token: str
    authored_locator: str
    resource_id: str


@dataclass(frozen=True, slots=True)
class ResolvedEmbeddedResource:
    resource_id: str
    role: Literal["linked_resource", "bibliography"]
    media_type: str
    size_bytes: int
    sha256: str
    content: bytes


@dataclass(frozen=True, slots=True)
class ResolvedCitationItem:
    citation_key: str
    record_id: str
    record_sha256: str
    presentation: str


@dataclass(frozen=True, slots=True)
class ResolvedCitation:
    source_start: int
    source_end: int
    source_slice_sha256: str
    authored_token: str
    form: Literal["narrative", "parenthetical"]
    cluster_id: str
    items: tuple[ResolvedCitationItem, ...]
    cached_result: str


@dataclass(frozen=True, slots=True)
class ResolvedDocument:
    authored_markdown: str
    targets: tuple[ResolvedDocumentTarget, ...]
    references: tuple[ResolvedReference, ...]
    resource_occurrences: tuple[ResolvedResourceOccurrence, ...]
    citations: tuple[ResolvedCitation, ...]
    resources: tuple[ResolvedEmbeddedResource, ...]


@dataclass(frozen=True, slots=True)
class NumberingTarget:
    source_start: int
    source_end: int
    kind: TargetKind
    enabled: bool
    target_id: str | None
    derived_number: str | None
    materialization: TargetMaterialization | None

    @property
    def occurrence_key(self) -> tuple[int, int, TargetKind]:
        return (self.source_start, self.source_end, self.kind)


@dataclass(frozen=True, slots=True)
class ResolvedNumberingPlan:
    heading_definitions: tuple[HeadingDefinition, ...]
    heading_instances: tuple[HeadingInstance, ...]
    targets: tuple[NumberingTarget, ...]


@dataclass(frozen=True, slots=True)
class ResolvedDocumentEnvelope:
    input_id: str
    source_sha256: str
    plan_sha256: str
    document: ResolvedDocument


@dataclass(frozen=True, slots=True)
class NumberingExportPlanEnvelope:
    input_id: str
    source_sha256: str
    plan_sha256: str
    plan: ResolvedNumberingPlan


@dataclass(frozen=True, slots=True)
class ResolvedNumberingPort:
    document_envelope: ResolvedDocumentEnvelope
    plan_envelope: NumberingExportPlanEnvelope

    @property
    def input_id(self) -> str:
        return self.document_envelope.input_id

    @property
    def source_sha256(self) -> str:
        return self.document_envelope.source_sha256

    @property
    def plan_sha256(self) -> str:
        return self.document_envelope.plan_sha256

    @property
    def document(self) -> ResolvedDocument:
        return self.document_envelope.document

    @property
    def plan(self) -> ResolvedNumberingPlan:
        return self.plan_envelope.plan


def load_resolved_numbering_port(
    neutral_document_path: str | Path,
    numbering_export_plan_path: str | Path,
) -> ResolvedNumberingPort:
    """Read and validate one immutable dual-input snapshot."""

    from docwen_core.models._resolved_numbering_validation import load_resolved_numbering_port

    return load_resolved_numbering_port(Path(neutral_document_path), Path(numbering_export_plan_path))


def load_resolved_numbering_bytes(
    neutral_document_bytes: bytes,
    numbering_export_plan_bytes: bytes,
) -> ResolvedNumberingPort:
    """Validate already-read bytes; intended for deterministic adapters/tests."""

    from docwen_core.models._resolved_numbering_validation import load_resolved_numbering_bytes

    return load_resolved_numbering_bytes(neutral_document_bytes, numbering_export_plan_bytes)


def canonicalize_numbering_plan(plan: object) -> bytes:
    """Return the restricted RFC 8785 canonical bytes used by ``plan_sha256``."""

    from docwen_core.models._resolved_numbering_validation import canonicalize_numbering_plan

    return canonicalize_numbering_plan(plan)


__all__ = [
    "MAX_RESOLVED_DOCUMENT_EMBEDDED_BYTES",
    "MAX_RESOLVED_NUMBERING_RESOURCE_BYTES",
    "NUMBERING_EXPORT_PLAN_MEDIA_TYPE",
    "NUMBERING_EXPORT_PLAN_SCHEMA",
    "NUMBERING_EXPORT_PLAN_SCHEMA_ID",
    "RESOLVED_DOCUMENT_MEDIA_TYPE",
    "RESOLVED_DOCUMENT_SCHEMA",
    "RESOLVED_DOCUMENT_SCHEMA_ID",
    "CaptionMaterialization",
    "CaptionNumberFormat",
    "HeadingCounterSegment",
    "HeadingDefinition",
    "HeadingDisplaySegment",
    "HeadingInstance",
    "HeadingLevelDefinition",
    "HeadingListMaterialization",
    "HeadingLiteralSegment",
    "HeadingNumberFormat",
    "HeadingStart",
    "NumberingExportPlanEnvelope",
    "NumberingTarget",
    "ResolvedCitation",
    "ResolvedCitationItem",
    "ResolvedDocument",
    "ResolvedDocumentEnvelope",
    "ResolvedDocumentTarget",
    "ResolvedEmbeddedResource",
    "ResolvedNumberingPlan",
    "ResolvedNumberingPort",
    "ResolvedNumberingPortError",
    "ResolvedReference",
    "ResolvedResourceOccurrence",
    "TargetKind",
    "TargetMaterialization",
    "canonicalize_numbering_plan",
    "load_resolved_numbering_bytes",
    "load_resolved_numbering_port",
]
