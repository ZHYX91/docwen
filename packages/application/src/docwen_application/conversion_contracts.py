"""Typed application conversion requests, plans, outcomes, and public capability shapes."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Literal

from docwen_core.models import (
    ArtifactBundle,
    ConversionDiagnostic,
    ConversionErrorInfo,
    ConversionMetrics,
)

MARKDOWN_TO_DOCX_CAPABILITY_ID = "convert.markdown.to_docx"
MARKDOWN_TO_XLSX_CAPABILITY_ID = "convert.markdown.to_xlsx"
DOCX_TO_MARKDOWN_CAPABILITY_ID = "convert.docx.to_markdown"
XLSX_TO_MARKDOWN_CAPABILITY_ID = "convert.xlsx.to_markdown"
PDF_TO_MARKDOWN_CAPABILITY_ID = "convert.pdf.to_markdown"
OFD_TO_MARKDOWN_CAPABILITY_ID = "convert.ofd.to_markdown"
XPS_TO_MARKDOWN_CAPABILITY_ID = "convert.xps.to_markdown"
TIFF_TO_MARKDOWN_CAPABILITY_ID = "convert.tiff.to_markdown"
XLSX_TO_CSV_CAPABILITY_ID = "convert.xlsx.to_csv"
MARKDOWN_TABLES_TO_CSV_CAPABILITY_ID = "convert.markdown_tables.to_csv"
PDF_TO_PNG_CAPABILITY_ID = "render.pdf.to_png"
PDF_SPLIT_EVERY_PAGE_CAPABILITY_ID = "split.pdf.every_page"
PNG_TO_OCR_MARKDOWN_CAPABILITY_ID = "convert.png.to_ocr_markdown"
TIFF_FRAMES_TO_PNG_CAPABILITY_ID = "convert.tiff_frames.to_png"
MARKDOWN_VALIDATE_CAPABILITY_ID = "validate.markdown"
MARKDOWN_NUMBERING_CAPABILITY_ID = "transform.markdown.heading_numbering"
PDF_MERGE_CAPABILITY_ID = "merge.pdf.documents"
PDF_SPLIT_CUSTOM_CAPABILITY_ID = "split.pdf.partition"
XLSX_MERGE_TABLES_CAPABILITY_ID = "merge.xlsx.tables"
IMAGES_MERGE_TO_TIFF_CAPABILITY_ID = "merge.images.to_tiff"
MARKDOWN_MEDIA_TYPE = "text/markdown"
DOCX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
XLSX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
PDF_MEDIA_TYPE = "application/pdf"
OFD_MEDIA_TYPE = "application/vnd.ofd"
XPS_MEDIA_TYPE = "application/vnd.ms-xpsdocument"
CSV_MEDIA_TYPE = "text/csv"
PNG_MEDIA_TYPE = "image/png"
TIFF_MEDIA_TYPE = "image/tiff"
JPEG_MEDIA_TYPE = "image/jpeg"
GIF_MEDIA_TYPE = "image/gif"
BMP_MEDIA_TYPE = "image/bmp"
WEBP_MEDIA_TYPE = "image/webp"
JSON_MEDIA_TYPE = "application/json"


class ConversionServiceError(ValueError):
    """Stable application-service failure ready for machine error projection."""

    def __init__(
        self,
        category: str,
        code: str,
        message: str,
        *,
        retryable: bool = False,
        details: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(message)
        self.category = category
        self.code = code
        self.retryable = retryable
        self.details = dict(details or {})

    def to_dict(self) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "category": self.category,
            "code": self.code,
            "message": str(self),
            "retryable": self.retryable,
        }
        if self.details:
            payload["details"] = dict(self.details)
        return payload


@dataclass(frozen=True, slots=True)
class LocalInputHandle:
    input_id: str
    path: str
    media_type: str
    size_bytes: int
    sha256: str
    kind: Literal["document", "resource"]
    role: Literal[
        "source",
        "linked_resource",
        "bibliography",
        "citation_style",
        "neutral_document",
        "numbering_export_plan",
    ]
    logical_path: str


@dataclass(frozen=True, slots=True)
class StagingOutputTarget:
    staging_root: str
    staging_policy: str = "require_empty"


@dataclass(frozen=True, slots=True)
class ConversionPlanRequest:
    capability_id: str
    inputs: tuple[LocalInputHandle, ...]
    output: StagingOutputTarget
    options: dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True, slots=True)
class OutputShape:
    cardinality: Literal["one", "many"]
    artifact_kinds: tuple[str, ...]
    relation_types: tuple[str, ...]
    relation_payloads: tuple[Literal["page_fragment", "page_resource"], ...] = ()
    atomic_bundle: bool = True

    def to_dict(self) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "cardinality": self.cardinality,
            "artifact_kinds": list(self.artifact_kinds),
            "relation_types": list(self.relation_types),
            "atomic_bundle": self.atomic_bundle,
        }
        if self.relation_payloads:
            payload["relation_payloads"] = list(self.relation_payloads)
        return payload


@dataclass(frozen=True, slots=True)
class InputSlot:
    role: Literal[
        "source",
        "linked_resource",
        "bibliography",
        "citation_style",
        "neutral_document",
        "numbering_export_plan",
    ]
    kind: Literal["document", "resource"]
    media_types: tuple[str, ...]
    min_items: int
    max_items: int | None = None

    def to_dict(self) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "role": self.role,
            "kind": self.kind,
            "media_types": list(self.media_types),
            "min_items": self.min_items,
        }
        if self.max_items is not None:
            payload["max_items"] = self.max_items
        return payload


@dataclass(frozen=True, slots=True)
class InputShape:
    slots: tuple[InputSlot, ...]
    undeclared_roles: Literal["reject"] = "reject"

    def to_dict(self) -> dict[str, Any]:
        return {
            "slots": [slot.to_dict() for slot in self.slots],
            "undeclared_roles": self.undeclared_roles,
        }


@dataclass(frozen=True, slots=True)
class MachineCapability:
    capability_id: str
    operation: str
    input_shape: InputShape
    output_media_types: tuple[str, ...]
    output_shape: OutputShape
    options_schema: dict[str, Any]
    availability: str
    dependencies: tuple[dict[str, Any], ...] = ()
    limitations: tuple[dict[str, Any], ...] = ()
    optimization_id: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "capability_id": self.capability_id,
            "operation": self.operation,
            "input_shape": self.input_shape.to_dict(),
            "output_media_types": list(self.output_media_types),
            "output_shape": self.output_shape.to_dict(),
            "options_schema": dict(self.options_schema),
            "availability": self.availability,
            "dependencies": [dict(item) for item in self.dependencies],
            "limitations": [dict(item) for item in self.limitations],
            **({"optimization_id": self.optimization_id} if self.optimization_id is not None else {}),
        }


@dataclass(frozen=True, slots=True)
class ConversionPlan:
    plan_id: str
    capability_id: str
    effective_options: dict[str, Any]
    output_shape: OutputShape
    warnings: tuple[dict[str, Any], ...] = ()
    limitations: tuple[dict[str, Any], ...] = ()
    requires_confirmation: bool = False

    def to_dict(self) -> dict[str, Any]:
        return {
            "plan_id": self.plan_id,
            "capability_id": self.capability_id,
            "effective_options": dict(self.effective_options),
            "output_shape": self.output_shape.to_dict(),
            "warnings": [dict(item) for item in self.warnings],
            "limitations": [dict(item) for item in self.limitations],
            "requires_confirmation": self.requires_confirmation,
        }


@dataclass(frozen=True, slots=True)
class ConversionTaskOutcome:
    task_id: str
    state: Literal["completed", "failed", "cancelled"]
    bundle: ArtifactBundle | None
    diagnostics: tuple[ConversionDiagnostic, ...]
    metrics: ConversionMetrics
    error: ConversionErrorInfo | None = None
