"""Plan-first application service shared by interactive and machine entry points."""

from __future__ import annotations

import hashlib
import logging
import os
import sys
import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Literal, Protocol
from uuid import uuid4

from docwen_application.bundle_mapping import (
    BundleMappingError,
    BundleProfile,
    build_bundle_draft,
    validate_physical_page_diagnostics,
)
from docwen_application.ports.runtime import ArtifactBundleCommitPort
from docwen_core.detection import FileAdmissionPathError
from docwen_core.docx_styles import SHIPPED_STYLE_LOCALES
from docwen_core.formats import (
    CATEGORY_DOCUMENT,
    CATEGORY_IMAGE,
    CATEGORY_LAYOUT,
    CATEGORY_MARKDOWN,
    CATEGORY_SPREADSHEET,
)
from docwen_core.models import (
    ArtifactBundle,
    ArtifactBundleValidationError,
    ConversionDiagnostic,
    ConversionErrorInfo,
    ConversionManifestContext,
    ConversionManifestInput,
    ConversionMetrics,
    ConversionRequest,
    ConversionResult,
    FileRef,
    OutputManifestPolicy,
    OutputPolicy,
    validate_artifact_bundle_draft,
)
from docwen_core.models.resolved_numbering import (
    NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
    RESOLVED_DOCUMENT_MEDIA_TYPE,
    ResolvedNumberingPortError,
    load_resolved_numbering_bytes,
)
from docwen_core.paths import filesystem_path
from docwen_core.semantic_bibliography import SEMANTIC_BIBLIOGRAPHY_MEDIA_TYPE

logger = logging.getLogger(__name__)

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
_HASH_CHUNK_BYTES = 1024 * 1024
_DOCUMENT_SEMANTICS_MACHINE_LIMITATIONS: tuple[dict[str, Any], ...] = (
    {
        "severity": "warning",
        "code": "document_semantics.citation_processor_unavailable",
        "message": (
            "DocWen does not run a CSL citation processor or accept citation_style inputs in Machine v1; "
            "Markdown citation keys remain literal."
        ),
    },
    {
        "severity": "warning",
        "code": "document_semantics.v1_scope",
        "message": (
            "Document semantics v1 excludes CSL processing, composite or range citation semantics, "
            "custom citation display, and PDF semantic round trips."
        ),
    },
)
_RESOLVED_DOCUMENT_MACHINE_LIMITATIONS: tuple[dict[str, Any], ...] = (
    {
        "severity": "warning",
        "code": "resolved_document.provider_owned_semantics",
        "message": (
            "DocWen consumes already-resolved targets, citations, resources, and numbering facts; it does not "
            "scan a Workspace, run a citation resolver, or infer numbering from authored text."
        ),
    },
)
_PHYSICAL_PAGE_OCR_LIMITATIONS: tuple[dict[str, Any], ...] = (
    {
        "severity": "warning",
        "code": "physical_page_ocr.best_effort",
        "message": (
            "When OCR is enabled it is best effort: every physical page or frame retains an ordered fragment and "
            "typed status even when recognition is blank or unavailable."
        ),
    },
    {
        "severity": "warning",
        "code": "physical_page_ocr.consumer_owned_import",
        "message": (
            "The Bundle reports page and resource facts only; Node layout, basenames, and import strategy remain "
            "consumer-owned."
        ),
    },
)


class _ExecutionController(Protocol):
    @property
    def has_runtime(self) -> bool: ...

    def describe_runtime_capabilities(self) -> dict[str, Any]: ...

    def prepare_execution_cancellation(self, request: Any, *, batch: bool = False) -> object: ...

    def release_execution_cancellation(self, task_id: str, reservation: object) -> None: ...

    def execute_single(self, request: Any) -> Any: ...

    def execute_aggregate(self, request: Any, action_name: str) -> Any: ...

    def cancel(self, task_id: str) -> None: ...


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


_SINGLE_DOCUMENT_SHAPE = OutputShape(
    cardinality="one",
    artifact_kinds=("document",),
    relation_types=(),
)

_DOCUMENT_WITH_ROUND_TRIP_SIDECAR_SHAPE = OutputShape(
    cardinality="many",
    artifact_kinds=("document", "resource"),
    relation_types=("resource_of",),
)

_DOCUMENT_WITH_RESOURCES_SHAPE = OutputShape(
    cardinality="many",
    artifact_kinds=("document", "fragment", "resource"),
    relation_types=("fragment_of", "resource_of"),
)

_PHYSICAL_PAGE_OCR_SHAPE = OutputShape(
    cardinality="many",
    artifact_kinds=("document", "fragment", "resource"),
    relation_types=("fragment_of", "resource_of"),
    relation_payloads=("page_fragment", "page_resource"),
)

_IMAGE_TO_OCR_MARKDOWN_SHAPE = OutputShape(
    cardinality="many",
    artifact_kinds=("document", "fragment", "resource"),
    relation_types=("fragment_of", "resource_of", "derived_from"),
)

_WORKSHEET_RESOURCES_SHAPE = OutputShape(
    cardinality="many",
    artifact_kinds=("resource",),
    relation_types=(),
)

_PAGE_IMAGES_SHAPE = OutputShape(
    cardinality="many",
    artifact_kinds=("resource",),
    relation_types=(),
)

_SECTION_DOCUMENTS_SHAPE = OutputShape(
    cardinality="many",
    artifact_kinds=("document",),
    relation_types=(),
)

_SINGLE_RESOURCE_SHAPE = OutputShape(
    cardinality="one",
    artifact_kinds=("resource",),
    relation_types=(),
)


def _strict_options(properties: dict[str, Any], *, required: tuple[str, ...] = ()) -> dict[str, Any]:
    return {
        "$schema": "https://json-schema.org/draft/2020-12/schema",
        "type": "object",
        "properties": properties,
        "required": list(required),
        "additionalProperties": False,
    }


_MARKDOWN_TO_DOCX_OPTIONS = _strict_options(
    {
        "locale": {
            "type": "string",
            "enum": list(SHIPPED_STYLE_LOCALES),
            "default": "zh_CN",
        },
        "template_name": {
            "type": "string",
            "pattern": r"^template\.docx\.[0-9a-f]{64}$",
            "x-docwen-resource-kind": "templates",
            "x-docwen-resource-target": "docx",
        },
        "heading_merge_mode": {
            "type": "string",
            "enum": ["punct_required", "never", "always"],
            "default": "punct_required",
        },
    }
)

_MARKDOWN_TO_XLSX_OPTIONS = _strict_options(
    {
        "template_name": {
            "type": "string",
            "pattern": r"^template\.xlsx\.[0-9a-f]{64}$",
            "x-docwen-resource-kind": "templates",
            "x-docwen-resource-target": "xlsx",
        }
    }
)

_XLSX_TO_MARKDOWN_OPTIONS = _strict_options(
    {
        "to_md_keep_images": {"type": "boolean", "default": True},
        "to_md_enable_ocr": {"type": "boolean", "default": False},
        "ocr_language": {
            "type": "string",
            "enum": ["auto", "chinese", "chinese_cht", "english", "japanese", "korean", "latin", "cyrillic"],
            "default": "auto",
        },
        "image_mode": {
            "type": "string",
            "enum": ["file", "base64", "embed", "omit"],
            "default": "file",
        },
        "ocr_placement": {
            "type": "string",
            "enum": ["image_md", "main_md"],
            "default": "image_md",
        },
        "image_link_style": {
            "type": "string",
            "enum": ["wiki_embed", "wiki_link", "markdown_embed", "markdown_link"],
            "default": "wiki_embed",
        },
        "table_merge_strategy": {
            "type": "string",
            "enum": ["fill", "empty", "marker"],
            "default": "fill",
        },
    }
)

_DOCX_TO_MARKDOWN_OPTIONS = _strict_options(
    {
        **{
            key: value
            for key, value in _XLSX_TO_MARKDOWN_OPTIONS["properties"].items()
            if key not in {"to_md_enable_ocr", "to_md_keep_images"}
        },
        "recognize_text": {"type": "boolean", "default": False},
        "preserve_resources": {"type": "boolean", "default": True},
        "ocr_placement": {
            "type": "string",
            "enum": ["image_md", "main_md"],
            "default": "main_md",
        },
        "remove_numbering": {"type": "boolean", "default": True},
        "add_numbering": {"type": "boolean", "default": False},
        "numbering_scheme": {
            "type": "string",
            "default": "gongwen_standard",
            "x-docwen-resource-kind": "numbering-schemes",
        },
    }
)

_PHYSICAL_PAGE_OCR_COMMON_PROPERTIES: dict[str, Any] = {
    "recognize_text": {"type": "boolean", "default": False},
    "preserve_resources": {"type": "boolean", "default": True},
    "ocr_language": {
        "type": "string",
        "enum": ["auto", "chinese", "chinese_cht", "english", "japanese", "korean", "latin", "cyrillic"],
        "default": "auto",
    },
}

_FIXED_LAYOUT_TO_MARKDOWN_OPTIONS = _strict_options(
    {
        **_PHYSICAL_PAGE_OCR_COMMON_PROPERTIES,
        "image_mode": {"type": "string", "enum": ["file"], "default": "file"},
        "render_dpi": {"type": "integer", "minimum": 72, "maximum": 600, "default": 200},
    }
)

_TIFF_TO_MARKDOWN_OPTIONS = _strict_options(dict(_PHYSICAL_PAGE_OCR_COMMON_PROPERTIES))

_MARKDOWN_VALIDATE_OPTIONS = _strict_options(
    {
        "enable_symbol_pairing": {"type": "boolean", "default": True},
        "enable_symbol_correction": {"type": "boolean", "default": True},
        "enable_typos_rule": {"type": "boolean", "default": True},
        "enable_sensitive_word": {"type": "boolean", "default": True},
        "skip_code_blocks": {"type": "boolean", "default": True},
        "skip_quote_blocks": {"type": "boolean", "default": False},
    }
)

_MARKDOWN_NUMBERING_OPTIONS = _strict_options(
    {
        "remove_numbering": {"type": "boolean", "default": True},
        "add_numbering": {"type": "boolean", "default": False},
        "numbering_scheme": {
            "type": "string",
            "default": "gongwen_standard",
            "x-docwen-resource-kind": "numbering-schemes",
        },
    }
)

_PDF_SPLIT_PARTITION_OPTIONS = _strict_options(
    {
        "pages": {
            "type": "array",
            "minItems": 1,
            "uniqueItems": True,
            "items": {"type": "integer", "minimum": 1},
        }
    },
    required=("pages",),
)

_XLSX_MERGE_TABLES_OPTIONS = _strict_options(
    {
        "merge_mode": {
            "type": "string",
            "enum": ["row", "col", "cell"],
            "default": "cell",
        },
        "offset_range": {
            "type": "integer",
            "minimum": 0,
            "maximum": 50,
            "default": 10,
        },
    }
)

_IMAGES_MERGE_TO_TIFF_OPTIONS = _strict_options(
    {
        "mode": {"type": "string", "enum": ["smart", "rgb", "RGB"], "default": "smart"},
        "keep_alpha": {"type": "boolean", "default": True},
    }
)


@dataclass(frozen=True, slots=True)
class _CapabilityBinding:
    capability_id: str
    input_media_type: str
    input_format: str
    input_category: str
    target_format: str
    output_media_type: str
    runtime_route_id: str
    operation: str = "convert"
    output_shape: OutputShape = _SINGLE_DOCUMENT_SHAPE
    bundle_profile: BundleProfile = "single_document"
    action_name: str = ""
    effective_options: dict[str, Any] = field(default_factory=dict)
    options_schema: dict[str, Any] = field(
        default_factory=lambda: {
            "$schema": "https://json-schema.org/draft/2020-12/schema",
            "type": "object",
            "properties": {},
            "additionalProperties": False,
        }
    )
    limitations: tuple[dict[str, Any], ...] = ()
    required_dependency_ids: tuple[str, ...] = ()
    dependency_ids: tuple[str, ...] = ()
    project_runtime_limitations: bool = True
    accepted_input_media_types: tuple[str, ...] = ()
    input_cardinality: Literal["one", "many"] = "one"
    minimum_inputs: int = 1

    @property
    def input_shape(self) -> InputShape:
        if self.capability_id == MARKDOWN_TO_DOCX_CAPABILITY_ID:
            return InputShape(
                slots=(
                    InputSlot(
                        role="neutral_document",
                        kind="document",
                        media_types=(RESOLVED_DOCUMENT_MEDIA_TYPE,),
                        min_items=1,
                        max_items=1,
                    ),
                    InputSlot(
                        role="numbering_export_plan",
                        kind="resource",
                        media_types=(NUMBERING_EXPORT_PLAN_MEDIA_TYPE,),
                        min_items=1,
                        max_items=1,
                    ),
                )
            )
        source_kind: Literal["document", "resource"] = (
            "document" if self.input_category in {CATEGORY_DOCUMENT, CATEGORY_MARKDOWN} else "resource"
        )
        accepted = (self.input_media_type, *self.accepted_input_media_types)
        source = InputSlot(
            role="source",
            kind=source_kind,
            media_types=accepted,
            min_items=self.minimum_inputs,
            max_items=1 if self.input_cardinality == "one" else None,
        )
        return InputShape(slots=(source,))


_CAPABILITY_BINDINGS = (
    _CapabilityBinding(
        capability_id=MARKDOWN_TO_DOCX_CAPABILITY_ID,
        input_media_type=RESOLVED_DOCUMENT_MEDIA_TYPE,
        input_format="markdown",
        input_category=CATEGORY_MARKDOWN,
        target_format="docx",
        output_media_type=DOCX_MEDIA_TYPE,
        runtime_route_id="docwen_plugin_markdown:markdown:docx:convert",
        output_shape=_DOCUMENT_WITH_ROUND_TRIP_SIDECAR_SHAPE,
        bundle_profile="document_with_round_trip_sidecar",
        options_schema=_MARKDOWN_TO_DOCX_OPTIONS,
        limitations=_RESOLVED_DOCUMENT_MACHINE_LIMITATIONS,
    ),
    _CapabilityBinding(
        capability_id=MARKDOWN_TO_XLSX_CAPABILITY_ID,
        input_media_type=MARKDOWN_MEDIA_TYPE,
        input_format="markdown",
        input_category=CATEGORY_MARKDOWN,
        target_format="xlsx",
        output_media_type=XLSX_MEDIA_TYPE,
        runtime_route_id="docwen_plugin_markdown:markdown:xlsx:convert",
        options_schema=_MARKDOWN_TO_XLSX_OPTIONS,
    ),
    _CapabilityBinding(
        capability_id=DOCX_TO_MARKDOWN_CAPABILITY_ID,
        input_media_type=DOCX_MEDIA_TYPE,
        input_format="docx",
        input_category=CATEGORY_DOCUMENT,
        target_format="md",
        output_media_type=MARKDOWN_MEDIA_TYPE,
        runtime_route_id="docwen_plugin_document:docx:md:convert",
        output_shape=_DOCUMENT_WITH_RESOURCES_SHAPE,
        bundle_profile="document_with_resources",
        options_schema=_DOCX_TO_MARKDOWN_OPTIONS,
        limitations=_DOCUMENT_SEMANTICS_MACHINE_LIMITATIONS,
    ),
    _CapabilityBinding(
        capability_id=PDF_TO_MARKDOWN_CAPABILITY_ID,
        input_media_type=PDF_MEDIA_TYPE,
        input_format="pdf",
        input_category=CATEGORY_LAYOUT,
        target_format="md",
        output_media_type=MARKDOWN_MEDIA_TYPE,
        runtime_route_id="docwen_plugin_layout:pdf:md:convert",
        output_shape=_PHYSICAL_PAGE_OCR_SHAPE,
        bundle_profile="physical_page_ocr",
        options_schema=_FIXED_LAYOUT_TO_MARKDOWN_OPTIONS,
        limitations=_PHYSICAL_PAGE_OCR_LIMITATIONS,
    ),
    _CapabilityBinding(
        capability_id=OFD_TO_MARKDOWN_CAPABILITY_ID,
        input_media_type=OFD_MEDIA_TYPE,
        input_format="ofd",
        input_category=CATEGORY_LAYOUT,
        target_format="md",
        output_media_type=MARKDOWN_MEDIA_TYPE,
        runtime_route_id="docwen_plugin_layout:ofd:md:convert",
        output_shape=_PHYSICAL_PAGE_OCR_SHAPE,
        bundle_profile="physical_page_ocr",
        options_schema=_FIXED_LAYOUT_TO_MARKDOWN_OPTIONS,
        limitations=_PHYSICAL_PAGE_OCR_LIMITATIONS,
    ),
    _CapabilityBinding(
        capability_id=XPS_TO_MARKDOWN_CAPABILITY_ID,
        input_media_type=XPS_MEDIA_TYPE,
        input_format="xps",
        input_category=CATEGORY_LAYOUT,
        target_format="md",
        output_media_type=MARKDOWN_MEDIA_TYPE,
        runtime_route_id="docwen_plugin_layout:xps:md:convert",
        output_shape=_PHYSICAL_PAGE_OCR_SHAPE,
        bundle_profile="physical_page_ocr",
        options_schema=_FIXED_LAYOUT_TO_MARKDOWN_OPTIONS,
        limitations=_PHYSICAL_PAGE_OCR_LIMITATIONS,
    ),
    _CapabilityBinding(
        capability_id=XLSX_TO_CSV_CAPABILITY_ID,
        input_media_type=XLSX_MEDIA_TYPE,
        input_format="xlsx",
        input_category=CATEGORY_SPREADSHEET,
        target_format="csv",
        output_media_type=CSV_MEDIA_TYPE,
        runtime_route_id="docwen_plugin_spreadsheet:xlsx:csv:convert",
        output_shape=_WORKSHEET_RESOURCES_SHAPE,
        bundle_profile="worksheet_resources",
    ),
    _CapabilityBinding(
        capability_id=XLSX_TO_MARKDOWN_CAPABILITY_ID,
        input_media_type=XLSX_MEDIA_TYPE,
        input_format="xlsx",
        input_category=CATEGORY_SPREADSHEET,
        target_format="md",
        output_media_type=MARKDOWN_MEDIA_TYPE,
        runtime_route_id="docwen_plugin_spreadsheet:xlsx:md:convert",
        output_shape=_DOCUMENT_WITH_RESOURCES_SHAPE,
        bundle_profile="document_with_resources",
        options_schema=_XLSX_TO_MARKDOWN_OPTIONS,
    ),
    _CapabilityBinding(
        capability_id=PDF_TO_PNG_CAPABILITY_ID,
        input_media_type=PDF_MEDIA_TYPE,
        input_format="pdf",
        input_category=CATEGORY_LAYOUT,
        target_format="png",
        output_media_type=PNG_MEDIA_TYPE,
        runtime_route_id="docwen_plugin_layout:pdf:png:convert",
        operation="render",
        output_shape=_PAGE_IMAGES_SHAPE,
        bundle_profile="page_images",
        effective_options={"render_dpi": 150},
    ),
    _CapabilityBinding(
        capability_id=PDF_SPLIT_EVERY_PAGE_CAPABILITY_ID,
        input_media_type=PDF_MEDIA_TYPE,
        input_format="pdf",
        input_category=CATEGORY_LAYOUT,
        target_format="pdf",
        output_media_type=PDF_MEDIA_TYPE,
        runtime_route_id="docwen_plugin_layout:pdf:pdf:split_pdf",
        operation="transform",
        output_shape=_SECTION_DOCUMENTS_SHAPE,
        bundle_profile="section_documents",
        action_name="split_pdf",
        effective_options={"split_mode": "every_page"},
    ),
    _CapabilityBinding(
        capability_id=PNG_TO_OCR_MARKDOWN_CAPABILITY_ID,
        input_media_type=PNG_MEDIA_TYPE,
        input_format="png",
        input_category=CATEGORY_IMAGE,
        target_format="md",
        output_media_type=MARKDOWN_MEDIA_TYPE,
        runtime_route_id="docwen_plugin_image:image:md:convert",
        output_shape=_IMAGE_TO_OCR_MARKDOWN_SHAPE,
        bundle_profile="image_to_markdown",
        effective_options={
            "image_mode": "file",
            "to_md_keep_images": True,
            "to_md_enable_ocr": True,
            "ocr_placement": "image_md",
        },
        required_dependency_ids=("python.rapidocr",),
        dependency_ids=("python.pillow", "python.rapidocr"),
        project_runtime_limitations=False,
    ),
    _CapabilityBinding(
        capability_id=MARKDOWN_TABLES_TO_CSV_CAPABILITY_ID,
        input_media_type=MARKDOWN_MEDIA_TYPE,
        input_format="markdown",
        input_category=CATEGORY_MARKDOWN,
        target_format="csv",
        output_media_type=CSV_MEDIA_TYPE,
        runtime_route_id="docwen_plugin_markdown:markdown:csv:convert",
        output_shape=_WORKSHEET_RESOURCES_SHAPE,
        bundle_profile="table_resources",
    ),
    _CapabilityBinding(
        capability_id=TIFF_FRAMES_TO_PNG_CAPABILITY_ID,
        input_media_type=TIFF_MEDIA_TYPE,
        input_format="tif",
        input_category=CATEGORY_IMAGE,
        target_format="png",
        output_media_type=PNG_MEDIA_TYPE,
        runtime_route_id="docwen_plugin_image:image:png:convert",
        output_shape=_PAGE_IMAGES_SHAPE,
        bundle_profile="frame_images",
        dependency_ids=("python.pillow",),
        project_runtime_limitations=False,
    ),
    _CapabilityBinding(
        capability_id=TIFF_TO_MARKDOWN_CAPABILITY_ID,
        input_media_type=TIFF_MEDIA_TYPE,
        input_format="tif",
        input_category=CATEGORY_IMAGE,
        target_format="md",
        output_media_type=MARKDOWN_MEDIA_TYPE,
        runtime_route_id="docwen_plugin_image:image:md:convert",
        output_shape=_PHYSICAL_PAGE_OCR_SHAPE,
        bundle_profile="physical_page_ocr",
        options_schema=_TIFF_TO_MARKDOWN_OPTIONS,
        limitations=_PHYSICAL_PAGE_OCR_LIMITATIONS,
        dependency_ids=("python.pillow", "python.rapidocr"),
        project_runtime_limitations=False,
    ),
    _CapabilityBinding(
        capability_id=MARKDOWN_VALIDATE_CAPABILITY_ID,
        input_media_type=MARKDOWN_MEDIA_TYPE,
        input_format="markdown",
        input_category=CATEGORY_MARKDOWN,
        target_format="markdown",
        output_media_type=JSON_MEDIA_TYPE,
        runtime_route_id="docwen_plugin_proofread:markdown:markdown:validate",
        operation="validate",
        output_shape=_SINGLE_RESOURCE_SHAPE,
        bundle_profile="report_resource",
        action_name="validate",
        options_schema=_MARKDOWN_VALIDATE_OPTIONS,
    ),
    _CapabilityBinding(
        capability_id=MARKDOWN_NUMBERING_CAPABILITY_ID,
        input_media_type=MARKDOWN_MEDIA_TYPE,
        input_format="markdown",
        input_category=CATEGORY_MARKDOWN,
        target_format="md",
        output_media_type=MARKDOWN_MEDIA_TYPE,
        runtime_route_id="docwen_plugin_markdown:markdown:md:process_md_numbering",
        operation="transform",
        action_name="process_md_numbering",
        options_schema=_MARKDOWN_NUMBERING_OPTIONS,
    ),
    _CapabilityBinding(
        capability_id=PDF_MERGE_CAPABILITY_ID,
        input_media_type=PDF_MEDIA_TYPE,
        input_format="pdf",
        input_category=CATEGORY_LAYOUT,
        target_format="pdf",
        output_media_type=PDF_MEDIA_TYPE,
        runtime_route_id="docwen_plugin_layout:pdf:pdf:merge_pdfs",
        operation="merge",
        action_name="merge_pdfs",
        input_cardinality="many",
        minimum_inputs=2,
    ),
    _CapabilityBinding(
        capability_id=PDF_SPLIT_CUSTOM_CAPABILITY_ID,
        input_media_type=PDF_MEDIA_TYPE,
        input_format="pdf",
        input_category=CATEGORY_LAYOUT,
        target_format="pdf",
        output_media_type=PDF_MEDIA_TYPE,
        runtime_route_id="docwen_plugin_layout:pdf:pdf:split_pdf",
        operation="transform",
        output_shape=_SECTION_DOCUMENTS_SHAPE,
        bundle_profile="partition_documents",
        action_name="split_pdf",
        effective_options={"split_mode": "custom"},
        options_schema=_PDF_SPLIT_PARTITION_OPTIONS,
    ),
    _CapabilityBinding(
        capability_id=XLSX_MERGE_TABLES_CAPABILITY_ID,
        input_media_type=XLSX_MEDIA_TYPE,
        input_format="spreadsheet",
        input_category=CATEGORY_SPREADSHEET,
        target_format="xlsx",
        output_media_type=XLSX_MEDIA_TYPE,
        runtime_route_id="docwen_plugin_spreadsheet:spreadsheet:xlsx:merge_tables",
        operation="merge",
        action_name="merge_tables",
        input_cardinality="many",
        minimum_inputs=2,
        options_schema=_XLSX_MERGE_TABLES_OPTIONS,
    ),
    _CapabilityBinding(
        capability_id=IMAGES_MERGE_TO_TIFF_CAPABILITY_ID,
        input_media_type=PNG_MEDIA_TYPE,
        input_format="image",
        input_category=CATEGORY_IMAGE,
        target_format="tif",
        output_media_type=TIFF_MEDIA_TYPE,
        runtime_route_id="docwen_plugin_image:image:tif:merge_images_to_tiff",
        operation="merge",
        output_shape=_SINGLE_RESOURCE_SHAPE,
        bundle_profile="image_resource",
        action_name="merge_images_to_tiff",
        accepted_input_media_types=(JPEG_MEDIA_TYPE, GIF_MEDIA_TYPE, BMP_MEDIA_TYPE, TIFF_MEDIA_TYPE, WEBP_MEDIA_TYPE),
        input_cardinality="many",
        minimum_inputs=2,
        options_schema=_IMAGES_MERGE_TO_TIFF_OPTIONS,
    ),
)
_CAPABILITY_BY_ID = {binding.capability_id: binding for binding in _CAPABILITY_BINDINGS}


@dataclass(frozen=True, slots=True)
class _RuntimeCapabilityState:
    availability: Literal["available", "limited", "unavailable"]
    dependencies: tuple[dict[str, Any], ...] = ()
    limitations: tuple[dict[str, Any], ...] = ()


@dataclass(frozen=True, slots=True)
class _RuntimeDiscovery:
    gates: dict[str, bool]
    routes: dict[str, dict[str, Any]]
    error: dict[str, Any] | None = None


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


@dataclass(slots=True)
class _PlanRecord:
    public: ConversionPlan
    request: ConversionPlanRequest
    binding: _CapabilityBinding


@dataclass(slots=True)
class _TaskRecord:
    request: ConversionRequest
    public_options: dict[str, Any]
    staging_root: str
    reservation: object
    binding: _CapabilityBinding
    state: str = "accepted"


class ConversionService:
    """Own capability discovery, immutable planning, execution, and cancellation."""

    def __init__(self, controller: _ExecutionController, bundle_committer: ArtifactBundleCommitPort) -> None:
        self._controller = controller
        self._bundle_committer = bundle_committer
        self._lock = threading.Lock()
        self._plans: dict[str, _PlanRecord] = {}
        self._tasks: dict[str, _TaskRecord] = {}

    def list_capabilities(self) -> tuple[MachineCapability, ...]:
        discovery = self._discover_runtime()
        capabilities: list[MachineCapability] = []
        for binding in _CAPABILITY_BINDINGS:
            state = self._runtime_capability_state(binding, discovery)
            capabilities.append(
                MachineCapability(
                    capability_id=binding.capability_id,
                    operation=binding.operation,
                    input_shape=binding.input_shape,
                    output_media_types=(binding.output_media_type,),
                    output_shape=binding.output_shape,
                    options_schema=dict(binding.options_schema),
                    availability=state.availability,
                    dependencies=state.dependencies,
                    limitations=(*binding.limitations, *state.limitations),
                )
            )
        return tuple(capabilities)

    def plan(self, request: ConversionPlanRequest) -> ConversionPlan:
        binding = self._validate_plan_request(request)
        effective_options = self._effective_options(request.options, binding)
        plan = ConversionPlan(
            plan_id=f"plan.{uuid4().hex}",
            capability_id=request.capability_id,
            effective_options=effective_options,
            output_shape=binding.output_shape,
            limitations=binding.limitations,
        )
        with self._lock:
            self._plans[plan.plan_id] = _PlanRecord(public=plan, request=request, binding=binding)
        return plan

    def accept(self, plan_id: str, task_id: str | None = None) -> str:
        with self._lock:
            record = self._plans.pop(plan_id, None)
        if record is None:
            raise ConversionServiceError(
                "conflict", "plan_not_found", f"plan is unknown or already consumed: {plan_id}"
            )

        task_id = task_id or f"task.{uuid4().hex}"
        self._validate_identifier(task_id, field_name="task_id")
        binding = self._validate_plan_request(record.request)
        if binding != record.binding:
            raise ConversionServiceError("conflict", "capability_changed", "capability changed after planning")
        effective_options = self._effective_options(record.request.options, binding)
        if effective_options != record.public.effective_options:
            raise ConversionServiceError(
                "conflict",
                "plan_options_changed",
                "capability options changed after planning",
            )
        conversion_request = self._build_conversion_request(
            task_id,
            record.request,
            binding,
            effective_options,
        )
        reservation = self._controller.prepare_execution_cancellation(conversion_request, batch=False)
        task = _TaskRecord(
            request=conversion_request,
            public_options=dict(effective_options),
            staging_root=record.request.output.staging_root,
            reservation=reservation,
            binding=binding,
        )
        with self._lock:
            if task_id in self._tasks:
                self._controller.release_execution_cancellation(task_id, reservation)
                raise ConversionServiceError("conflict", "task_id_in_use", f"task id is already in use: {task_id}")
            self._tasks[task_id] = task
        return task_id

    def execute_accepted(self, task_id: str) -> ConversionTaskOutcome:
        with self._lock:
            task = self._tasks.get(task_id)
            if task is None:
                raise ConversionServiceError("conflict", "task_not_found", f"task is not accepted: {task_id}")
            if task.state != "accepted":
                raise ConversionServiceError("conflict", "task_not_accepted", f"task cannot execute from {task.state}")
            task.state = "running"

        try:
            self._validate_execution_snapshot(task.request, task.staging_root)
            try:
                if task.binding.input_cardinality == "many":
                    raw_result = self._controller.execute_aggregate(task.request, task.binding.action_name)
                else:
                    raw_result = self._controller.execute_single(task.request)
            except FileAdmissionPathError as exc:
                raise ConversionServiceError("security", exc.error_type, str(exc)) from exc
            if not isinstance(raw_result, ConversionResult):
                raise ConversionServiceError(
                    "internal",
                    "invalid_runtime_result",
                    "runtime returned an unsupported result type",
                )
            outcome = self._outcome_from_result(task, raw_result)
            return outcome
        finally:
            self._controller.release_execution_cancellation(task_id, task.reservation)
            with self._lock:
                task.state = "terminal"

    def cancel(self, task_id: str) -> str:
        with self._lock:
            task = self._tasks.get(task_id)
            if task is None:
                return "not_found"
            if task.state == "terminal":
                return "already_terminal"
        self._controller.cancel(task_id)
        return "cancellation_requested"

    def _outcome_from_result(self, task: _TaskRecord, result: ConversionResult) -> ConversionTaskOutcome:
        if result.task_id != task.request.request_id:
            raise ConversionServiceError(
                "internal",
                "runtime_task_mismatch",
                "runtime result belongs to a different task",
            )
        if not result.success:
            self._discard_result_artifacts(task.staging_root, result)
            bound_diagnostics = sorted(
                {diagnostic.artifact_id for diagnostic in result.diagnostics if diagnostic.artifact_id is not None}
            )
            if bound_diagnostics:
                raise ConversionServiceError(
                    "conversion_failed",
                    "dangling_diagnostic_artifact",
                    "failed or cancelled runtime result cannot bind diagnostics to output artifacts",
                    details={"artifact_ids": bound_diagnostics},
                )
            state: Literal["failed", "cancelled"] = (
                "cancelled" if result.error is not None and result.error.error_type == "cancelled" else "failed"
            )
            return ConversionTaskOutcome(
                task_id=result.task_id,
                state=state,
                bundle=None,
                diagnostics=tuple(result.diagnostics),
                metrics=result.metrics,
                error=result.error,
            )

        try:
            draft = build_bundle_draft(
                profile=task.binding.bundle_profile,
                output_media_type=task.binding.output_media_type,
                artifacts=result.artifacts,
            )
        except BundleMappingError as exc:
            self._discard_result_artifacts(task.staging_root, result)
            raise ConversionServiceError(
                exc.category,
                exc.code,
                str(exc),
                details=exc.details,
            ) from exc
        try:
            validate_artifact_bundle_draft(draft)
        except ArtifactBundleValidationError as exc:
            self._discard_result_artifacts(task.staging_root, result)
            raise ConversionServiceError(
                "conversion_failed",
                exc.code,
                str(exc),
            ) from exc
        artifact_ids = {artifact.artifact_id for artifact in draft.artifacts}
        dangling_diagnostics = sorted(
            {
                diagnostic.artifact_id
                for diagnostic in result.diagnostics
                if diagnostic.artifact_id is not None and diagnostic.artifact_id not in artifact_ids
            }
        )
        if dangling_diagnostics:
            self._discard_result_artifacts(task.staging_root, result)
            raise ConversionServiceError(
                "conversion_failed",
                "dangling_diagnostic_artifact",
                "runtime diagnostic references an artifact outside the output bundle",
                details={"artifact_ids": dangling_diagnostics},
            )
        if task.binding.capability_id == DOCX_TO_MARKDOWN_CAPABILITY_ID:
            expected_recognition = task.public_options.get("recognize_text")
            expected_resources = task.public_options.get("preserve_resources")
            expected_placement = task.public_options.get("ocr_placement")
            image_relations = [
                relation for relation in draft.relations if relation.type == "resource_of" and relation.role == "image"
            ]
            ocr_relations = [
                relation
                for relation in draft.relations
                if relation.type == "fragment_of" and relation.role == "ocr_text"
            ]
            invalid_public_contract = (
                not isinstance(expected_recognition, bool)
                or not isinstance(expected_resources, bool)
                or expected_placement not in {"image_md", "main_md"}
            )
            producer_drift = (not expected_resources and bool(image_relations)) or (
                (not expected_recognition or expected_placement == "main_md") and bool(ocr_relations)
            )
            if invalid_public_contract or producer_drift:
                self._discard_result_artifacts(task.staging_root, result)
                raise ConversionServiceError(
                    "internal",
                    "document_fidelity_option_mismatch",
                    "DOCX producer artifacts do not match the accepted fidelity options",
                    details={
                        "image_resource_count": len(image_relations),
                        "ocr_fragment_count": len(ocr_relations),
                    },
                )
        if task.binding.bundle_profile == "physical_page_ocr":
            try:
                validate_physical_page_diagnostics(draft, result.diagnostics)
            except BundleMappingError as exc:
                self._discard_result_artifacts(task.staging_root, result)
                raise ConversionServiceError(
                    exc.category,
                    exc.code,
                    str(exc),
                    details=exc.details,
                ) from exc
            preferred = [artifact for artifact in result.artifacts if artifact.is_primary]
            expected_ocr = task.public_options.get("recognize_text")
            expected_images = task.public_options.get("preserve_resources")
            if (
                len(preferred) != 1
                or not isinstance(expected_ocr, bool)
                or not isinstance(expected_images, bool)
                or preferred[0].metadata.get("ocr_enabled") is not expected_ocr
                or preferred[0].metadata.get("keep_images") is not expected_images
            ):
                self._discard_result_artifacts(task.staging_root, result)
                raise ConversionServiceError(
                    "internal",
                    "physical_page_option_mismatch",
                    "physical-page producer modes do not match the accepted Machine options",
                )
        try:
            bundle = self._bundle_committer.commit(
                task_id=result.task_id,
                staging_root=task.staging_root,
                draft=draft,
            )
        except Exception as exc:
            logger.exception("Artifact Bundle commit rejected for task %s", result.task_id)
            self._discard_result_artifacts(task.staging_root, result)
            raise ConversionServiceError(
                "security",
                "bundle_commit_failed",
                "runtime rejected the output bundle",
            ) from exc
        return ConversionTaskOutcome(
            task_id=result.task_id,
            state="completed",
            bundle=bundle,
            diagnostics=tuple(result.diagnostics),
            metrics=result.metrics,
        )

    def _discard_result_artifacts(self, staging_root: str, result: ConversionResult) -> None:
        paths = [artifact.staging_path for artifact in result.artifacts]
        if not paths:
            return
        try:
            self._bundle_committer.discard(staging_root=staging_root, artifact_paths=paths)
        except Exception:
            # Preserve the authoritative task failure. The machine adapter never
            # projects rejected paths as deliverables even if local cleanup is denied.
            return

    def _validate_plan_request(self, request: ConversionPlanRequest) -> _CapabilityBinding:
        binding = _CAPABILITY_BY_ID.get(request.capability_id)
        if binding is None:
            raise ConversionServiceError(
                "unsupported",
                "capability_not_found",
                f"unsupported capability: {request.capability_id}",
            )
        if not self._controller.has_runtime:
            raise ConversionServiceError("unavailable", "runtime_unavailable", "conversion runtime is unavailable")
        runtime_state = self._runtime_capability_state(binding, self._discover_runtime())
        if runtime_state.availability == "unavailable":
            missing = [
                dependency["dependency_id"]
                for dependency in runtime_state.dependencies
                if dependency["required"] and not dependency["available"]
            ]
            raise ConversionServiceError(
                "unavailable",
                "capability_unavailable",
                f"capability is unavailable in the active runtime: {binding.capability_id}",
                details={"missing_required_dependencies": missing},
            )
        input_ids = [handle.input_id for handle in request.inputs]
        if len(input_ids) != len(set(input_ids)):
            raise ConversionServiceError(
                "invalid_request",
                "duplicate_input_id",
                "task inputs must use unique input identifiers",
            )
        for handle in request.inputs:
            self._validate_input_logical_path(handle.logical_path)
        logical_paths = [handle.logical_path for handle in request.inputs]
        if len(logical_paths) != len(set(logical_paths)):
            raise ConversionServiceError(
                "invalid_request",
                "duplicate_input_logical_path",
                "task inputs must use unique logical paths",
            )
        slots = {slot.role: slot for slot in binding.input_shape.slots}
        for handle in request.inputs:
            slot = slots.get(handle.role)
            if slot is None:
                raise ConversionServiceError(
                    "invalid_request",
                    "undeclared_input_role",
                    f"capability does not declare input role: {handle.role}",
                )
        for handle in request.inputs:
            slot = slots[handle.role]
            if handle.kind != slot.kind:
                raise ConversionServiceError(
                    "invalid_request",
                    "input_slot_kind_mismatch",
                    f"input role {handle.role} requires kind {slot.kind}",
                )
        for handle in request.inputs:
            slot = slots[handle.role]
            if handle.media_type not in slot.media_types:
                raise ConversionServiceError(
                    "unsupported",
                    "input_slot_media_type_mismatch",
                    f"input role {handle.role} does not accept {handle.media_type}",
                )
        if binding.capability_id == MARKDOWN_TO_DOCX_CAPABILITY_ID:
            neutral_count = sum(handle.role == "neutral_document" for handle in request.inputs)
            plan_count = sum(handle.role == "numbering_export_plan" for handle in request.inputs)
            if neutral_count == 0:
                raise ConversionServiceError(
                    "invalid_request",
                    "docwen.resolved_document.missing",
                    "resolved-document input is required",
                )
            if plan_count == 0:
                raise ConversionServiceError(
                    "invalid_request",
                    "docwen.numbering_export_plan.missing",
                    "numbering-export-plan input is required",
                )
            if neutral_count != 1:
                raise ConversionServiceError(
                    "invalid_request",
                    "docwen.resolved_document.invalid",
                    "resolved-document input must occur exactly once",
                )
            if plan_count != 1:
                raise ConversionServiceError(
                    "invalid_request",
                    "docwen.numbering_export_plan.invalid",
                    "numbering-export-plan input must occur exactly once",
                )
        for slot in binding.input_shape.slots:
            count = sum(handle.role == slot.role for handle in request.inputs)
            if count < slot.min_items or (slot.max_items is not None and count > slot.max_items):
                raise ConversionServiceError(
                    "invalid_request",
                    "input_slot_cardinality_mismatch",
                    f"input role {slot.role} has invalid cardinality",
                )
        self._effective_options(request.options, binding)
        if request.output.staging_policy != "require_empty":
            raise ConversionServiceError(
                "invalid_request",
                "unsupported_staging_policy",
                "staging_policy must be require_empty",
            )
        for handle in request.inputs:
            self._validate_input(handle)
        if binding.capability_id == MARKDOWN_TO_DOCX_CAPABILITY_ID:
            self._validate_resolved_numbering_inputs(request)
        self._validate_empty_staging_root(request.output.staging_root)
        return binding

    @staticmethod
    def _validate_resolved_numbering_inputs(request: ConversionPlanRequest) -> None:
        neutral = next(handle for handle in request.inputs if handle.role == "neutral_document")
        plan = next(handle for handle in request.inputs if handle.role == "numbering_export_plan")
        try:
            neutral_bytes = Path(neutral.path).read_bytes()
        except OSError as exc:
            raise ConversionServiceError(
                "security",
                "docwen.resolved_document.invalid",
                "resolved-document input cannot be read",
            ) from exc
        try:
            plan_bytes = Path(plan.path).read_bytes()
        except OSError as exc:
            raise ConversionServiceError(
                "security",
                "docwen.numbering_export_plan.invalid",
                "numbering-export-plan input cannot be read",
            ) from exc
        if len(neutral_bytes) != neutral.size_bytes or hashlib.sha256(neutral_bytes).hexdigest() != neutral.sha256:
            raise ConversionServiceError(
                "conflict",
                "docwen.resolved_document.invalid",
                "resolved-document changed during admission",
            )
        if len(plan_bytes) != plan.size_bytes or hashlib.sha256(plan_bytes).hexdigest() != plan.sha256:
            raise ConversionServiceError(
                "conflict",
                "docwen.numbering_export_plan.invalid",
                "numbering-export-plan changed during admission",
            )
        try:
            load_resolved_numbering_bytes(neutral_bytes, plan_bytes)
        except ResolvedNumberingPortError as exc:
            category = (
                "unsupported"
                if exc.code == "docwen.numbering_export_plan.unsupported_materialization"
                else "invalid_request"
            )
            raise ConversionServiceError(category, exc.code, str(exc)) from exc

    @classmethod
    def _effective_options(
        cls,
        provided: dict[str, Any],
        binding: _CapabilityBinding,
    ) -> dict[str, Any]:
        schema = binding.options_schema
        properties = schema.get("properties", {})
        if not isinstance(properties, dict):
            raise ConversionServiceError(
                "internal",
                "capability_options_schema_invalid",
                "capability has an invalid options schema",
            )
        unknown = sorted(set(provided) - set(properties))
        if unknown:
            raise ConversionServiceError(
                "invalid_request",
                "unsupported_options",
                "capability does not accept one or more caller-defined options",
                details={"option_keys": unknown},
            )
        effective = {
            key: property_schema["default"]
            for key, property_schema in properties.items()
            if isinstance(property_schema, dict) and "default" in property_schema
        }
        effective.update(binding.effective_options)
        effective.update(provided)
        missing = [key for key in schema.get("required", []) if key not in effective]
        if missing:
            raise ConversionServiceError(
                "invalid_request",
                "required_options_missing",
                "capability requires one or more options",
                details={"option_keys": sorted(missing)},
            )
        for key, value in effective.items():
            property_schema = properties.get(key)
            if property_schema is None:
                continue
            cls._validate_option_value(key, value, property_schema)
        return effective

    @staticmethod
    def _validate_option_value(key: str, value: Any, schema: dict[str, Any]) -> None:
        raw_expected_type = schema.get("type")
        expected_type = raw_expected_type if isinstance(raw_expected_type, str) else ""
        valid_type = {
            "boolean": isinstance(value, bool),
            "string": isinstance(value, str),
            "integer": isinstance(value, int) and not isinstance(value, bool),
            "number": isinstance(value, (int, float)) and not isinstance(value, bool),
            "array": isinstance(value, list),
            "object": isinstance(value, dict),
        }.get(expected_type, True)
        invalid = not valid_type
        if not invalid and "enum" in schema:
            invalid = value not in schema["enum"]
        if not invalid and isinstance(value, (int, float)) and not isinstance(value, bool):
            minimum = schema.get("minimum")
            maximum = schema.get("maximum")
            invalid = (minimum is not None and value < minimum) or (maximum is not None and value > maximum)
        if not invalid and isinstance(value, list) and isinstance(schema.get("items"), dict):
            minimum_items = schema.get("minItems")
            if isinstance(minimum_items, int) and len(value) < minimum_items:
                invalid = True
            if schema.get("uniqueItems") is True and any(value[index] in value[:index] for index in range(len(value))):
                invalid = True
            try:
                for index, item in enumerate(value):
                    ConversionService._validate_option_value(f"{key}[{index}]", item, schema["items"])
            except ConversionServiceError:
                invalid = True
        if invalid:
            raise ConversionServiceError(
                "invalid_request",
                "option_value_invalid",
                f"option has a value outside its capability contract: {key}",
                details={"option_key": key},
            )

    def _discover_runtime(self) -> _RuntimeDiscovery:
        if not self._controller.has_runtime:
            return _RuntimeDiscovery(
                gates={},
                routes={},
                error={
                    "severity": "error",
                    "code": "runtime_unavailable",
                    "message": "The conversion runtime is not configured.",
                },
            )

        try:
            description = self._controller.describe_runtime_capabilities()
            gates = {
                str(gate.get("id")): bool(gate.get("available"))
                for gate in description.get("gates", [])
                if isinstance(gate, dict) and isinstance(gate.get("id"), str)
            }
            routes = {
                str(candidate["id"]): candidate
                for source in description.get("sources", [])
                if isinstance(source, dict)
                for candidate in source.get("routes", [])
                if isinstance(candidate, dict) and isinstance(candidate.get("id"), str)
            }
        except Exception:
            return _RuntimeDiscovery(
                gates={},
                routes={},
                error={
                    "severity": "error",
                    "code": "runtime_discovery_failed",
                    "message": "The active runtime could not describe its capabilities.",
                },
            )
        return _RuntimeDiscovery(gates=gates, routes=routes)

    @staticmethod
    def _runtime_capability_state(
        binding: _CapabilityBinding,
        discovery: _RuntimeDiscovery,
    ) -> _RuntimeCapabilityState:
        if discovery.error is not None:
            return _RuntimeCapabilityState(
                availability="unavailable",
                limitations=(discovery.error,),
            )

        route = discovery.routes.get(binding.runtime_route_id)
        if route is None:
            return _RuntimeCapabilityState(
                availability="unavailable",
                limitations=(
                    {
                        "severity": "error",
                        "code": "runtime_route_missing",
                        "message": f"The active runtime does not expose route {binding.runtime_route_id}.",
                    },
                ),
            )

        required_ids = tuple(
            dict.fromkeys(
                (
                    *(str(item) for item in route.get("required_capabilities", [])),
                    *binding.required_dependency_ids,
                )
            )
        )
        optional_ids = tuple(
            str(item) for item in route.get("optional_capabilities", []) if str(item) not in required_ids
        )
        if binding.dependency_ids:
            required_ids = tuple(item for item in required_ids if item in binding.dependency_ids)
            optional_ids = tuple(item for item in optional_ids if item in binding.dependency_ids)
        dependencies = tuple(
            {
                "dependency_id": dependency_id,
                "required": required,
                "available": discovery.gates.get(dependency_id, False),
            }
            for required, dependency_ids in (
                (True, required_ids),
                (False, optional_ids),
            )
            for dependency_id in dependency_ids
        )
        route_limitations = (
            tuple(
                {
                    "severity": "warning",
                    "code": "runtime_route_limitation",
                    "message": str(message),
                }
                for message in route.get("limitations", [])
                if str(message).strip()
            )
            if binding.project_runtime_limitations
            else ()
        )
        route_available = bool(route.get("available")) and not any(
            dependency["required"] and not dependency["available"] for dependency in dependencies
        )
        if not route_available:
            availability: Literal["available", "limited", "unavailable"] = "unavailable"
        elif any(not dependency["required"] and not dependency["available"] for dependency in dependencies):
            availability = "limited"
        else:
            availability = "available"
        return _RuntimeCapabilityState(
            availability=availability,
            dependencies=dependencies,
            limitations=route_limitations,
        )

    def _validate_execution_snapshot(self, request: ConversionRequest, staging_root: str) -> None:
        for input_ref in request.input_refs:
            expected_sha = str(input_ref.metadata["machine_input_sha256"])
            expected_size = int(input_ref.metadata["machine_input_size_bytes"])
            path = filesystem_path(input_ref.path, force_extended=sys.platform == "win32")
            if self._path_traverses_link_or_junction(path):
                raise ConversionServiceError(
                    "security",
                    "input_is_link",
                    "input must not be a link or junction",
                )
            size_bytes, sha256 = self._file_integrity(path, code_prefix="input")
            if size_bytes != expected_size or sha256 != expected_sha:
                raise ConversionServiceError(
                    "conflict",
                    "input_changed_after_plan",
                    "input content changed after planning",
                )
        self._validate_empty_staging_root(staging_root)

    def _validate_input(self, handle: LocalInputHandle) -> None:
        self._validate_identifier(handle.input_id, field_name="input_id")
        path = Path(handle.path)
        if not path.is_absolute():
            raise ConversionServiceError("invalid_request", "input_path_not_absolute", "input path must be absolute")
        io_path = filesystem_path(path, force_extended=sys.platform == "win32")
        if self._path_traverses_link_or_junction(io_path):
            raise ConversionServiceError(
                "security",
                "input_is_link",
                "input must not be a link or junction",
            )
        size_bytes, sha256 = self._file_integrity(io_path, code_prefix="input")
        if size_bytes != handle.size_bytes or sha256 != handle.sha256:
            raise ConversionServiceError(
                "invalid_request",
                "input_integrity_mismatch",
                "declared input size or sha256 does not match the local file",
            )

    @staticmethod
    def _validate_input_logical_path(logical_path: str) -> None:
        segments = logical_path.split("/")
        if (
            not logical_path
            or len(logical_path) > 1024
            or logical_path.startswith("/")
            or "\\" in logical_path
            or "\x00" in logical_path
            or ":" in segments[0]
            or any(segment in {"", ".", ".."} for segment in segments)
        ):
            raise ConversionServiceError(
                "invalid_request",
                "invalid_input_logical_path",
                "input logical_path must be a normalized relative POSIX path",
            )

    @classmethod
    def _validate_empty_staging_root(cls, staging_root: str) -> None:
        path = Path(staging_root)
        if not path.is_absolute():
            raise ConversionServiceError(
                "invalid_request",
                "staging_root_not_absolute",
                "staging root must be absolute",
            )
        io_path = filesystem_path(path, force_extended=sys.platform == "win32")
        if cls._is_link_or_junction(io_path):
            raise ConversionServiceError(
                "security",
                "staging_root_is_link",
                "staging root must not be a link or junction",
            )
        try:
            if not io_path.is_dir():
                raise ConversionServiceError(
                    "invalid_request",
                    "staging_root_not_directory",
                    "staging root must be an existing directory",
                )
            if next(io_path.iterdir(), None) is not None:
                raise ConversionServiceError(
                    "conflict",
                    "staging_root_not_empty",
                    "staging root must be empty",
                )
        except OSError as exc:
            raise ConversionServiceError(
                "security",
                "staging_root_unreadable",
                "staging root cannot be inspected",
            ) from exc

    @staticmethod
    def _validate_identifier(value: str, *, field_name: str) -> None:
        if not value or len(value) > 128 or not value[0].isalnum():
            raise ConversionServiceError(
                "invalid_request",
                "invalid_identifier",
                f"{field_name} is not a valid identifier",
            )
        if any(not (character.isascii() and (character.isalnum() or character in "._:-")) for character in value):
            raise ConversionServiceError(
                "invalid_request",
                "invalid_identifier",
                f"{field_name} is not a valid identifier",
            )

    @staticmethod
    def _is_link_or_junction(path: Path) -> bool:
        if path.is_symlink():
            return True
        is_junction = getattr(path, "is_junction", None)
        return bool(is_junction()) if callable(is_junction) else False

    @classmethod
    def _path_traverses_link_or_junction(cls, path: Path) -> bool:
        current = path
        while True:
            if cls._is_link_or_junction(current):
                return True
            parent = current.parent
            if parent == current:
                return False
            current = parent

    @staticmethod
    def _file_integrity(path: Path, *, code_prefix: str) -> tuple[int, str]:
        if not path.is_file():
            raise ConversionServiceError(
                "invalid_request",
                f"{code_prefix}_not_regular_file",
                f"{code_prefix} must be an existing regular file",
            )
        digest = hashlib.sha256()
        size_bytes = 0
        try:
            with path.open("rb") as stream:
                while chunk := stream.read(_HASH_CHUNK_BYTES):
                    size_bytes += len(chunk)
                    digest.update(chunk)
        except OSError as exc:
            raise ConversionServiceError(
                "security",
                f"{code_prefix}_unreadable",
                f"{code_prefix} cannot be read",
            ) from exc
        return size_bytes, digest.hexdigest()

    @staticmethod
    def _build_conversion_request(
        task_id: str,
        request: ConversionPlanRequest,
        binding: _CapabilityBinding,
        effective_options: dict[str, Any],
    ) -> ConversionRequest:
        manifest_context = ConversionManifestContext(
            policy=OutputManifestPolicy(save_to_output=False, mask_input_path=True),
            inputs=tuple(
                ConversionManifestInput(
                    path=handle.path,
                    format=binding.input_format,
                    category=binding.input_category,
                )
                for handle in request.inputs
                if handle.role in {"source", "neutral_document"}
            ),
        )
        runtime_options = dict(effective_options)
        public_properties = binding.options_schema.get("properties", {})
        if isinstance(public_properties, dict) and {
            "recognize_text",
            "preserve_resources",
        }.issubset(public_properties):
            runtime_options["to_md_enable_ocr"] = runtime_options.pop("recognize_text")
            runtime_options["to_md_keep_images"] = runtime_options.pop("preserve_resources")

        return ConversionRequest(
            request_id=task_id,
            input_refs=[
                FileRef(
                    path=os.path.abspath(handle.path),
                    format=(binding.input_format if handle.role in {"source", "neutral_document"} else "resource"),
                    category=(binding.input_category if handle.role in {"source", "neutral_document"} else "other"),
                    size_bytes=handle.size_bytes,
                    input_kind=handle.kind,
                    input_role=handle.role,
                    logical_path=handle.logical_path,
                    media_type=handle.media_type,
                    metadata={
                        "machine_input_id": handle.input_id,
                        "machine_input_size_bytes": handle.size_bytes,
                        "machine_input_sha256": handle.sha256,
                    },
                )
                for handle in request.inputs
            ],
            target_format=binding.target_format,
            action_name=binding.action_name,
            options=runtime_options,
            output_policy=OutputPolicy(
                output_dir=os.path.abspath(request.output.staging_root),
                overwrite_mode="error",
                write_artifacts=True,
                open_after_done=False,
            ),
            # The controller captures the complete request-scoped config snapshot.
            # Manifest policy is carried independently so Machine tasks never
            # publish the legacy sidecar manifest into their staging bundle.
            config_snapshot={},
            manifest_context=manifest_context,
        )


__all__ = [
    "CSV_MEDIA_TYPE",
    "DOCX_MEDIA_TYPE",
    "DOCX_TO_MARKDOWN_CAPABILITY_ID",
    "JSON_MEDIA_TYPE",
    "MARKDOWN_MEDIA_TYPE",
    "MARKDOWN_NUMBERING_CAPABILITY_ID",
    "MARKDOWN_TABLES_TO_CSV_CAPABILITY_ID",
    "MARKDOWN_TO_DOCX_CAPABILITY_ID",
    "MARKDOWN_TO_XLSX_CAPABILITY_ID",
    "MARKDOWN_VALIDATE_CAPABILITY_ID",
    "OFD_MEDIA_TYPE",
    "OFD_TO_MARKDOWN_CAPABILITY_ID",
    "PDF_MEDIA_TYPE",
    "PDF_SPLIT_EVERY_PAGE_CAPABILITY_ID",
    "PDF_TO_MARKDOWN_CAPABILITY_ID",
    "PDF_TO_PNG_CAPABILITY_ID",
    "PNG_MEDIA_TYPE",
    "PNG_TO_OCR_MARKDOWN_CAPABILITY_ID",
    "SEMANTIC_BIBLIOGRAPHY_MEDIA_TYPE",
    "TIFF_FRAMES_TO_PNG_CAPABILITY_ID",
    "TIFF_MEDIA_TYPE",
    "TIFF_TO_MARKDOWN_CAPABILITY_ID",
    "XLSX_MEDIA_TYPE",
    "XLSX_TO_CSV_CAPABILITY_ID",
    "XLSX_TO_MARKDOWN_CAPABILITY_ID",
    "XPS_MEDIA_TYPE",
    "XPS_TO_MARKDOWN_CAPABILITY_ID",
    "ConversionPlan",
    "ConversionPlanRequest",
    "ConversionService",
    "ConversionServiceError",
    "ConversionTaskOutcome",
    "LocalInputHandle",
    "MachineCapability",
    "OutputShape",
    "StagingOutputTarget",
]
