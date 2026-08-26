"""Route manifest for docwen_plugin_document.

Declares all conversion routes this plugin handles.

Declared routes:
    ROUTE-DOC-001     docx → md, document → md  (implemented)
    ROUTE-DOCFMT-*    20 SmartConverter routes    (implemented via external office bridge)
"""

from __future__ import annotations

from docwen_core.models.manifest import PluginManifest, RouteCapabilityRule, RouteSpec

PLUGIN_ID = "docwen_plugin_document"
PLUGIN_VERSION = "0.1.0"

# ── Option schemas ─────────────────────────────────────────────────────

DOCX_TO_MD_OPTIONS_SCHEMA: dict = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "to_md_keep_images": {
            "type": "boolean",
            "default": True,
            "description": (
                "Extract embedded images from DOCX and save as files, referenced via Markdown image syntax."
            ),
            "x-docwen-status": "implemented",
        },
        "remove_numbering": {
            "type": "boolean",
            "default": True,
            "description": "Remove heading numbering from output.",
            "x-docwen-status": "implemented",
        },
        "add_numbering": {
            "type": "boolean",
            "default": False,
            "description": "Add heading numbering to output using the specified scheme.",
            "x-docwen-status": "implemented",
        },
        "numbering_scheme": {
            "type": "string",
            "enum": ["gongwen_standard", "hierarchical_standard", "hierarchical_h2_start", "legal_standard"],
            "default": "gongwen_standard",
            "description": "Numbering scheme ID when add_numbering is True.",
            "x-docwen-status": "implemented",
        },
        "to_md_enable_ocr": {
            "type": "boolean",
            "default": False,
            "description": "Run OCR on extracted images and append recognised text as a blockquote.",
            "x-docwen-status": "implemented",
        },
        "ocr_language": {
            "type": "string",
            "enum": ["auto", "chinese", "chinese_cht", "english", "japanese", "korean", "latin", "cyrillic"],
            "default": "auto",
            "description": "OCR language used for extracted images.",
            "x-docwen-status": "implemented",
        },
        "locale": {
            "type": "string",
            "description": "App-resolved locale used for YAML frontmatter labels and OCR context.",
            "x-docwen-status": "implemented",
        },
        "yaml_key_labels": {
            "type": "object",
            "additionalProperties": {"type": "string"},
            "default": {},
            "description": "App-resolved YAML frontmatter key labels supplied by the request boundary.",
            "x-docwen-status": "implemented",
        },
        "image_mode": {
            "type": "string",
            "enum": ["file", "base64", "embed", "omit"],
            "default": "file",
            "description": "Image storage mode: file=separate file, base64=inline data URI, embed=local placeholder, omit=no image output.",
            "x-docwen-status": "implemented",
        },
        "ocr_placement": {
            "type": "string",
            "enum": ["image_md", "main_md"],
            "default": "main_md",
            "description": "Where to place OCR-extracted text relative to the image.",
            "x-docwen-status": "implemented",
        },
        "image_link_style": {
            "type": "string",
            "enum": ["wiki_embed", "wiki_link", "markdown_embed", "markdown_link"],
            "default": "wiki_embed",
            "description": "Markdown image link style for generated references.",
            "x-docwen-status": "implemented",
        },
        "table_merge_strategy": {
            "type": "string",
            "enum": ["fill", "empty", "marker"],
            "default": "fill",
            "description": "How merged table cells are rendered in Markdown tables.",
            "x-docwen-status": "implemented",
        },
        "preserve_formatting": {
            "type": "boolean",
            "default": True,
            "description": "Preserve inline formatting markers in DOCX body paragraphs.",
            "x-docwen-status": "implemented",
        },
        "preserve_heading_formatting": {
            "type": "boolean",
            "default": False,
            "description": "Preserve inline formatting markers inside DOCX headings.",
            "x-docwen-status": "implemented",
        },
        "preserve_table_header_formatting": {
            "type": "boolean",
            "default": False,
            "description": "Preserve inline formatting markers in the first Markdown table row.",
            "x-docwen-status": "implemented",
        },
        "page_break_separator": {
            "type": "string",
            "enum": ["---", "***", "___", "ignore"],
            "default": "---",
            "description": "Markdown separator emitted for DOCX page breaks.",
            "x-docwen-status": "implemented",
        },
        "section_break_separator": {
            "type": "string",
            "enum": ["---", "***", "___", "ignore"],
            "default": "***",
            "description": "Markdown separator emitted for DOCX section breaks.",
            "x-docwen-status": "implemented",
        },
        "horizontal_rule_separator": {
            "type": "string",
            "enum": ["---", "***", "___", "ignore"],
            "default": "___",
            "description": "Markdown separator emitted for DOCX horizontal-rule border groups.",
            "x-docwen-status": "implemented",
        },
        "code_block_style_aliases": {
            "type": "array",
            "items": {"type": "string"},
            "default": [],
            "description": "Advanced DOCX→MD style override: paragraph style fragments treated as code blocks.",
            "x-docwen-status": "implemented",
        },
        "quote_style_aliases": {
            "type": "array",
            "items": {"type": "string"},
            "default": [],
            "description": "Advanced DOCX→MD style override: paragraph style names treated as level-1 quotes.",
            "x-docwen-status": "implemented",
        },
        "quote_generic_names": {
            "type": "array",
            "items": {"type": "string"},
            "default": [],
            "description": "Advanced DOCX→MD style override: generic quote style names detected without numeric suffixes.",
            "x-docwen-status": "implemented",
        },
    },
    "required": [],
}

# Shared empty options schema for routes without extra options.
_NOT_IMPL_OPTIONS_SCHEMA: dict = {
    "type": "object",
    "additionalProperties": False,
    "properties": {},
    "required": [],
}


# ── Implemented routes ─────────────────────────────────────────────────

ROUTE_DOCX_TO_MD = RouteSpec(
    source_format="docx",
    target_format="md",
    action_name="",
    label="DOCX → Markdown (Standard/Gongwen/Invoice)",
    options_schema=DOCX_TO_MD_OPTIONS_SCHEMA,
)

ROUTE_DOCUMENT_TO_MD = RouteSpec(
    source_format="document",
    target_format="md",
    action_name="",
    label="Document → Markdown (category-level alias for DOCX→MD)",
    options_schema=DOCX_TO_MD_OPTIONS_SCHEMA,
)

# ── SmartConverter document-format interconversion ───────────────────────

# fmt: off — keep the 20 SmartConverter routes compact and legible
ROUTE_DOCX_TO_DOC = RouteSpec(
    source_format="docx",
    target_format="doc",
    action_name="",
    label="DOCX → DOC (SmartConverter)",
    options_schema=_NOT_IMPL_OPTIONS_SCHEMA,
)
ROUTE_DOCX_TO_ODT = RouteSpec(
    source_format="docx",
    target_format="odt",
    action_name="",
    label="DOCX → ODT (SmartConverter)",
    options_schema=_NOT_IMPL_OPTIONS_SCHEMA,
)
ROUTE_DOCX_TO_RTF = RouteSpec(
    source_format="docx",
    target_format="rtf",
    action_name="",
    label="DOCX → RTF (SmartConverter)",
    options_schema=_NOT_IMPL_OPTIONS_SCHEMA,
)
ROUTE_DOCX_TO_WPS = RouteSpec(
    source_format="docx",
    target_format="wps",
    action_name="",
    label="DOCX → WPS (SmartConverter)",
    options_schema=_NOT_IMPL_OPTIONS_SCHEMA,
)
ROUTE_DOC_TO_DOCX = RouteSpec(
    source_format="doc",
    target_format="docx",
    action_name="",
    label="DOC → DOCX (SmartConverter)",
    options_schema=_NOT_IMPL_OPTIONS_SCHEMA,
)
ROUTE_DOC_TO_ODT = RouteSpec(
    source_format="doc",
    target_format="odt",
    action_name="",
    label="DOC → ODT (SmartConverter)",
    options_schema=_NOT_IMPL_OPTIONS_SCHEMA,
)
ROUTE_DOC_TO_RTF = RouteSpec(
    source_format="doc",
    target_format="rtf",
    action_name="",
    label="DOC → RTF (SmartConverter)",
    options_schema=_NOT_IMPL_OPTIONS_SCHEMA,
)
ROUTE_DOC_TO_WPS = RouteSpec(
    source_format="doc",
    target_format="wps",
    action_name="",
    label="DOC → WPS (SmartConverter)",
    options_schema=_NOT_IMPL_OPTIONS_SCHEMA,
)
ROUTE_ODT_TO_DOCX = RouteSpec(
    source_format="odt",
    target_format="docx",
    action_name="",
    label="ODT → DOCX (SmartConverter)",
    options_schema=_NOT_IMPL_OPTIONS_SCHEMA,
)
ROUTE_ODT_TO_DOC = RouteSpec(
    source_format="odt",
    target_format="doc",
    action_name="",
    label="ODT → DOC (SmartConverter)",
    options_schema=_NOT_IMPL_OPTIONS_SCHEMA,
)
ROUTE_ODT_TO_RTF = RouteSpec(
    source_format="odt",
    target_format="rtf",
    action_name="",
    label="ODT → RTF (SmartConverter)",
    options_schema=_NOT_IMPL_OPTIONS_SCHEMA,
)
ROUTE_ODT_TO_WPS = RouteSpec(
    source_format="odt",
    target_format="wps",
    action_name="",
    label="ODT → WPS (SmartConverter)",
    options_schema=_NOT_IMPL_OPTIONS_SCHEMA,
)
ROUTE_RTF_TO_DOCX = RouteSpec(
    source_format="rtf",
    target_format="docx",
    action_name="",
    label="RTF → DOCX (SmartConverter)",
    options_schema=_NOT_IMPL_OPTIONS_SCHEMA,
)
ROUTE_RTF_TO_DOC = RouteSpec(
    source_format="rtf",
    target_format="doc",
    action_name="",
    label="RTF → DOC (SmartConverter)",
    options_schema=_NOT_IMPL_OPTIONS_SCHEMA,
)
ROUTE_RTF_TO_ODT = RouteSpec(
    source_format="rtf",
    target_format="odt",
    action_name="",
    label="RTF → ODT (SmartConverter)",
    options_schema=_NOT_IMPL_OPTIONS_SCHEMA,
)
ROUTE_RTF_TO_WPS = RouteSpec(
    source_format="rtf",
    target_format="wps",
    action_name="",
    label="RTF → WPS (SmartConverter)",
    options_schema=_NOT_IMPL_OPTIONS_SCHEMA,
)
ROUTE_WPS_TO_DOCX = RouteSpec(
    source_format="wps",
    target_format="docx",
    action_name="",
    label="WPS → DOCX (SmartConverter)",
    options_schema=_NOT_IMPL_OPTIONS_SCHEMA,
)
ROUTE_WPS_TO_DOC = RouteSpec(
    source_format="wps",
    target_format="doc",
    action_name="",
    label="WPS → DOC (SmartConverter)",
    options_schema=_NOT_IMPL_OPTIONS_SCHEMA,
)
ROUTE_WPS_TO_ODT = RouteSpec(
    source_format="wps",
    target_format="odt",
    action_name="",
    label="WPS → ODT (SmartConverter)",
    options_schema=_NOT_IMPL_OPTIONS_SCHEMA,
)
ROUTE_WPS_TO_RTF = RouteSpec(
    source_format="wps",
    target_format="rtf",
    action_name="",
    label="WPS → RTF (SmartConverter)",
    options_schema=_NOT_IMPL_OPTIONS_SCHEMA,
)


# ── All routes ─────────────────────────────────────────────────────────

ALL_ROUTES: list[RouteSpec] = [
    # Implemented
    ROUTE_DOCX_TO_MD,
    ROUTE_DOCUMENT_TO_MD,
    # Implemented: SmartConverter document-format interconversion
    ROUTE_DOCX_TO_DOC,
    ROUTE_DOCX_TO_ODT,
    ROUTE_DOCX_TO_RTF,
    ROUTE_DOCX_TO_WPS,
    ROUTE_DOC_TO_DOCX,
    ROUTE_DOC_TO_ODT,
    ROUTE_DOC_TO_RTF,
    ROUTE_DOC_TO_WPS,
    ROUTE_ODT_TO_DOCX,
    ROUTE_ODT_TO_DOC,
    ROUTE_ODT_TO_RTF,
    ROUTE_ODT_TO_WPS,
    ROUTE_RTF_TO_DOCX,
    ROUTE_RTF_TO_DOC,
    ROUTE_RTF_TO_ODT,
    ROUTE_RTF_TO_WPS,
    ROUTE_WPS_TO_DOCX,
    ROUTE_WPS_TO_DOC,
    ROUTE_WPS_TO_ODT,
    ROUTE_WPS_TO_RTF,
]


def build_manifest() -> PluginManifest:
    """Build the plugin manifest for docwen_plugin_document."""
    return PluginManifest(
        plugin_id=PLUGIN_ID,
        name="DocWen Document Plugin",
        version=PLUGIN_VERSION,
        description=(
            "Converts document family formats (DOCX/DOC/ODT/RTF/WPS) to Markdown, "
            "with gongwen (official Chinese document) recognition, and document "
            "format interconversion via external office bridge."
        ),
        author="DocWen",
        routes=ALL_ROUTES,
        capability_rules=[
            RouteCapabilityRule(
                target_formats=("md",),
                required_capabilities=("python.docx",),
            ),
            RouteCapabilityRule(
                target_formats=("doc", "docx", "odt", "rtf", "wps"),
                required_capabilities=("external_office.word",),
                limitations=("output fidelity depends on the selected external Office backend",),
            ),
        ],
        extra={
            "third_party_deps": ["python-docx"],
            "external_bridge_routes": [
                "20 SmartConverter document-format routes (docx/doc/odt/rtf/wps ↔ docx/doc/odt/rtf/wps)",
            ],
        },
    )
