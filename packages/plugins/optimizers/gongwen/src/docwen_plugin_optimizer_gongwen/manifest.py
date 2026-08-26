"""Gongwen optimizer plugin manifest."""

from __future__ import annotations

from docwen_core.models.manifest import OptimizationResourceSpec, PluginManifest, RouteCapabilityRule, RouteSpec

PLUGIN_ID = "docwen_plugin_optimizer_gongwen"
PLUGIN_VERSION = "0.1.0"

GONGWEN_OPTIONS_SCHEMA: dict = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "remove_numbering": {
            "type": "boolean",
            "default": True,
            "description": "Remove heading numbering from output.",
        },
        "add_numbering": {
            "type": "boolean",
            "default": False,
            "description": "Add heading numbering to output using the specified scheme.",
        },
        "numbering_scheme": {
            "type": "string",
            "enum": ["gongwen_standard", "hierarchical_standard", "hierarchical_h2_start", "legal_standard"],
            "default": "gongwen_standard",
            "description": "Numbering scheme ID when add_numbering is True.",
        },
        "to_md_keep_images": {
            "type": "boolean",
            "default": True,
            "description": "Keep embedded image references and image artifacts in Markdown output.",
        },
        "image_mode": {
            "type": "string",
            "enum": ["file", "base64", "embed", "omit"],
            "default": "file",
            "description": "Image storage mode for Gongwen Markdown output.",
        },
        "image_link_style": {
            "type": "string",
            "enum": ["wiki_embed", "wiki_link", "markdown_embed", "markdown_link"],
            "default": "markdown_embed",
            "description": "Markdown link style for retained Gongwen images.",
        },
        "table_merge_strategy": {
            "type": "string",
            "enum": ["fill", "empty", "marker"],
            "default": "fill",
            "description": "Merged-cell projection strategy for retained tables.",
        },
        "to_md_enable_ocr": {
            "type": "boolean",
            "default": False,
            "description": "Run OCR on embedded images during Markdown export.",
        },
        "ocr_language": {
            "type": "string",
            "enum": ["auto", "chinese", "chinese_cht", "english", "japanese", "korean", "latin", "cyrillic"],
            "default": "auto",
            "description": "OCR language used for embedded images.",
        },
        "locale": {
            "type": "string",
            "description": "App-resolved locale used for OCR context; public-document fields remain fixed zh_CN.",
            "x-docwen-status": "implemented",
        },
    },
    "required": [],
}

ROUTE_DOCX_GONGWEN = RouteSpec(
    source_format="docx",
    target_format="md",
    action_name="gongwen",
    label="DOCX → Markdown (Chinese official document recognition)",
    options_schema=GONGWEN_OPTIONS_SCHEMA,
)

ALL_ROUTES: list[RouteSpec] = [
    ROUTE_DOCX_GONGWEN,
]


def build_manifest() -> PluginManifest:
    return PluginManifest(
        plugin_id=PLUGIN_ID,
        name="DocWen Gongwen Optimizer (zh_CN)",
        version=PLUGIN_VERSION,
        description=(
            "Recognises Chinese official document (公文) structure, extracting "
            "18 YAML metadata fields via three-round scoring and re-evaluation."
        ),
        author="DocWen",
        routes=ALL_ROUTES,
        capability_rules=[
            RouteCapabilityRule(
                required_capabilities=("python.docx",),
                optional_capabilities=("python.rapidocr",),
                limitations=("OCR options require the optional RapidOCR capability",),
            )
        ],
        optimization_resources=[
            OptimizationResourceSpec(
                id="gongwen",
                name="Chinese official-document optimization",
                action_name="gongwen",
            )
        ],
        extra={},
    )
