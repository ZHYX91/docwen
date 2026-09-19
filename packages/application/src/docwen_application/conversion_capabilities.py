"""Machine capability bindings and their supported option and artifact contracts."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Literal

from docwen_application.bundle_mapping import (
    BundleProfile,
)
from docwen_application.conversion_contracts import (
    BMP_MEDIA_TYPE,
    CSV_MEDIA_TYPE,
    DOCX_MEDIA_TYPE,
    DOCX_TO_MARKDOWN_CAPABILITY_ID,
    GIF_MEDIA_TYPE,
    IMAGES_MERGE_TO_TIFF_CAPABILITY_ID,
    JPEG_MEDIA_TYPE,
    JSON_MEDIA_TYPE,
    MARKDOWN_MEDIA_TYPE,
    MARKDOWN_NUMBERING_CAPABILITY_ID,
    MARKDOWN_TABLES_TO_CSV_CAPABILITY_ID,
    MARKDOWN_TO_DOCX_CAPABILITY_ID,
    MARKDOWN_TO_XLSX_CAPABILITY_ID,
    MARKDOWN_VALIDATE_CAPABILITY_ID,
    OFD_MEDIA_TYPE,
    OFD_TO_MARKDOWN_CAPABILITY_ID,
    PDF_MEDIA_TYPE,
    PDF_MERGE_CAPABILITY_ID,
    PDF_SPLIT_CUSTOM_CAPABILITY_ID,
    PDF_SPLIT_EVERY_PAGE_CAPABILITY_ID,
    PDF_TO_MARKDOWN_CAPABILITY_ID,
    PDF_TO_PNG_CAPABILITY_ID,
    PNG_MEDIA_TYPE,
    PNG_TO_OCR_MARKDOWN_CAPABILITY_ID,
    TIFF_FRAMES_TO_PNG_CAPABILITY_ID,
    TIFF_MEDIA_TYPE,
    TIFF_TO_MARKDOWN_CAPABILITY_ID,
    WEBP_MEDIA_TYPE,
    XLSX_MEDIA_TYPE,
    XLSX_MERGE_TABLES_CAPABILITY_ID,
    XLSX_TO_CSV_CAPABILITY_ID,
    XLSX_TO_MARKDOWN_CAPABILITY_ID,
    XPS_MEDIA_TYPE,
    XPS_TO_MARKDOWN_CAPABILITY_ID,
    InputShape,
    InputSlot,
    OutputShape,
)
from docwen_core.docx_styles import SHIPPED_STYLE_LOCALES
from docwen_core.formats import (
    CATEGORY_DOCUMENT,
    CATEGORY_IMAGE,
    CATEGORY_LAYOUT,
    CATEGORY_MARKDOWN,
    CATEGORY_SPREADSHEET,
)
from docwen_core.markdown_extensions import MARKDOWN_EXTENSIONS_OPTIONS_SCHEMA
from docwen_core.models.resolved_numbering import (
    NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
    RESOLVED_DOCUMENT_MEDIA_TYPE,
)


def _strict_options(properties: dict[str, Any], *, required: tuple[str, ...] = ()) -> dict[str, Any]:
    return {
        "$schema": "https://json-schema.org/draft/2020-12/schema",
        "type": "object",
        "properties": properties,
        "required": list(required),
        "additionalProperties": False,
    }


_PHYSICAL_PAGE_OCR_COMMON_PROPERTIES: dict[str, Any] = {
    "recognize_text": {"type": "boolean", "default": False},
    "preserve_resources": {"type": "boolean", "default": True},
    "ocr_language": {
        "type": "string",
        "enum": ["auto", "chinese", "chinese_cht", "english", "japanese", "korean", "latin", "cyrillic"],
        "default": "auto",
    },
}


_DOCUMENT_SEMANTICS_MACHINE_LIMITATIONS: tuple[dict[str, Any], ...] = (
    {
        "severity": "warning",
        "code": "document_semantics.citation_processor_unavailable",
        "message": (
            "DocWen does not run a CSL citation processor or accept citation_style inputs in Machine v2; "
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


_SINGLE_DOCUMENT_SHAPE = OutputShape(
    cardinality="one",
    artifact_kinds=("document",),
    relation_types=(),
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


_MARKDOWN_TO_DOCX_OPTIONS = _strict_options(
    {
        "markdown_extensions": MARKDOWN_EXTENSIONS_OPTIONS_SCHEMA,
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
        "markdown_extensions": MARKDOWN_EXTENSIONS_OPTIONS_SCHEMA,
        "template_name": {
            "type": "string",
            "pattern": r"^template\.xlsx\.[0-9a-f]{64}$",
            "x-docwen-resource-kind": "templates",
            "x-docwen-resource-target": "xlsx",
        },
    }
)


_XLSX_TO_MARKDOWN_OPTIONS = _strict_options(
    {
        "markdown_extensions": MARKDOWN_EXTENSIONS_OPTIONS_SCHEMA,
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
class CapabilityBinding:
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


CAPABILITY_BINDINGS = (
    CapabilityBinding(
        capability_id=MARKDOWN_TO_DOCX_CAPABILITY_ID,
        input_media_type=RESOLVED_DOCUMENT_MEDIA_TYPE,
        input_format="markdown",
        input_category=CATEGORY_MARKDOWN,
        target_format="docx",
        output_media_type=DOCX_MEDIA_TYPE,
        runtime_route_id="docwen_plugin_markdown:markdown:docx:convert",
        options_schema=_MARKDOWN_TO_DOCX_OPTIONS,
        limitations=_RESOLVED_DOCUMENT_MACHINE_LIMITATIONS,
    ),
    CapabilityBinding(
        capability_id=MARKDOWN_TO_XLSX_CAPABILITY_ID,
        input_media_type=MARKDOWN_MEDIA_TYPE,
        input_format="markdown",
        input_category=CATEGORY_MARKDOWN,
        target_format="xlsx",
        output_media_type=XLSX_MEDIA_TYPE,
        runtime_route_id="docwen_plugin_markdown:markdown:xlsx:convert",
        options_schema=_MARKDOWN_TO_XLSX_OPTIONS,
    ),
    CapabilityBinding(
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
    CapabilityBinding(
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
    CapabilityBinding(
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
    CapabilityBinding(
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
    CapabilityBinding(
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
    CapabilityBinding(
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
    CapabilityBinding(
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
    CapabilityBinding(
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
    CapabilityBinding(
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
    CapabilityBinding(
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
    CapabilityBinding(
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
    CapabilityBinding(
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
    CapabilityBinding(
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
    CapabilityBinding(
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
    CapabilityBinding(
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
    CapabilityBinding(
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
    CapabilityBinding(
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
    CapabilityBinding(
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


CAPABILITY_BY_ID = {binding.capability_id: binding for binding in CAPABILITY_BINDINGS}
