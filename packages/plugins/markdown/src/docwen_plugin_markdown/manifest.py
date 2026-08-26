"""Route manifest for docwen_plugin_markdown.

Declares every conversion and action route this plugin handles. Office-backed
routes use ``docwen_core.office_bridge``.

Declared routes (11 total):
    ROUTE-MD-DOCX-001  markdown -> docx   (implemented)
    ROUTE-MD-DOC-001   markdown -> doc    (Office bridge-backed)
    ROUTE-MD-ODT-001   markdown -> odt    (Office bridge-backed)
    ROUTE-MD-RTF-001   markdown -> rtf    (Office bridge-backed)
    ROUTE-MD-WPS-001   markdown -> wps    (Office bridge-backed)
    ROUTE-MD-PDF-001   markdown -> pdf    (Office bridge-backed)
    ROUTE-MD-XLSX-001  markdown -> xlsx   (implemented)
    ROUTE-MD-XLS-001   markdown -> xls    (Office bridge-backed)
    ROUTE-MD-ODS-001   markdown -> ods    (Office bridge-backed)
    ROUTE-MD-CSV-001   markdown -> csv    (implemented)
    ACT-MD-NUMBERING   process_md_numbering  (implemented)
"""

from __future__ import annotations

from docwen_core.docx_styles import SHIPPED_STYLE_LOCALES
from docwen_core.models.manifest import HonestyRoute, PluginManifest, RouteCapabilityRule, RouteSpec
from docwen_core.text.heading_merge import DEFAULT_HEADING_MERGE_PUNCTUATION

PLUGIN_ID = "docwen_plugin_markdown"
PLUGIN_VERSION = "0.1.0"

# ── Option schemas ─────────────────────────────────────────────────────

MD_TO_DOCX_OPTIONS_SCHEMA: dict = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "locale": {
            "type": "string",
            "enum": list(SHIPPED_STYLE_LOCALES),
            "description": "Request-owned locale for DOCX style names and defaults.",
        },
        "remove_numbering": {
            "type": "boolean",
            "default": True,
            "description": ("Remove heading numbering from the source Markdown before writing DOCX."),
        },
        "add_numbering": {
            "type": "boolean",
            "default": False,
            "description": "Add heading numbering to output DOCX headings.",
        },
        "numbering_scheme": {
            "type": "string",
            "enum": [
                "gongwen_standard",
                "hierarchical_standard",
                "hierarchical_h2_start",
                "legal_standard",
            ],
            "default": "hierarchical_standard",
            "description": "Numbering scheme when add_numbering is True.",
        },
        "heading_numbering_render_mode": {
            "type": "string",
            "enum": ["text", "word_native"],
            "default": "text",
            "description": (
                "Heading numbering output mode: text=concatenate into heading text, word_native=Word multi-level list."
            ),
        },
        "formatting_mode": {
            "type": "string",
            "enum": ["full", "minimal", "keep"],
            "default": "full",
            "description": (
                'Inline formatting mode. "full" preserves all bold/italic/'
                'strikethrough formatting. "minimal" strips all inline '
                "formatting, keeping only plain text, links, images, and formulas. "
                '"keep" renders Markdown inline markers as visible text.'
            ),
        },
        "heading_formatting_mode": {
            "type": "string",
            "enum": ["apply", "keep", "remove"],
            "default": "remove",
            "description": (
                'Formatting mode for headings. "apply" uses override_style '
                'to enable partial bold/italic within headings. "keep" '
                'preserves markdown formatting as-is. "remove" strips all '
                "inline formatting from headings."
            ),
        },
        "table_header_formatting_mode": {
            "type": "string",
            "enum": ["apply", "keep", "remove"],
            "default": "remove",
            "description": (
                'Formatting mode for Markdown table header cells. "apply" '
                'preserves supported inline formatting, "keep" preserves '
                'Markdown markers as visible text, and "remove" strips '
                "inline formatting from header cells."
            ),
        },
        "code_font": {
            "type": "string",
            "default": "Consolas",
            "description": "Monospace font for code spans and code blocks.",
        },
        "code_background_color": {
            "type": "string",
            "default": "E7E6E6",
            "description": "Background color (RRGGBB) for code spans and code blocks.",
        },
        "table_style_mode": {
            "type": "string",
            "enum": ["builtin", "custom"],
            "default": "builtin",
            "description": "Table style resolution mode.",
        },
        "builtin_style_key": {
            "type": "string",
            "enum": ["three_line_table", "table_grid"],
            "default": "three_line_table",
            "description": "Built-in table style key when table_style_mode=builtin.",
        },
        "custom_style_name": {
            "type": "string",
            "default": "",
            "description": "Custom table style name when table_style_mode=custom.",
        },
        "heading_merge_mode": {
            "type": "string",
            "enum": ["punct_required", "never", "always"],
            "default": "punct_required",
            "description": (
                "Strategy for merging headings with following body text. "
                '"punct_required" only merges when heading ends with '
                'punctuation. "always" merges unconditionally. '
                '"never" disables heading merging.'
            ),
        },
        "heading_merge_punctuation": {
            "type": "string",
            "default": DEFAULT_HEADING_MERGE_PUNCTUATION,
            "description": (
                "Characters that trigger heading/body merging when "
                "heading_merge_mode=punct_required. Whitespace is ignored "
                "and duplicate characters have no effect."
            ),
        },
        "template_name": {
            "type": "string",
            "pattern": r"^template\.docx\.[0-9a-f]{64}$",
            "description": (
                "Canonical DOCX template resource ID published by resources list templates. "
                "Names, filenames, paths, and blank values are rejected. "
                "If omitted, the built-in A4 template is used."
            ),
        },
        "hr_mapping": {
            "type": "string",
            "enum": ["", "dash", "asterisk", "underscore"],
            "default": "",
            "description": (
                "Horizontal rule rendering. Empty string uses Word built-in "
                'horizontal rule border. "dash", "asterisk", or '
                '"underscore" renders a centered text separator instead.'
            ),
        },
    },
    "required": [],
}

# The active resolved-document capability has a separate exact input shape
# and must never inherit the historical source-authoring controls above.
# Application publishes this closed schema for the exact-two Machine route;
# the plugin rechecks the same key set before touching either resource.
RESOLVED_V4_MD_TO_DOCX_OPTIONS_SCHEMA: dict = {
    "$schema": "https://json-schema.org/draft/2020-12/schema",
    "type": "object",
    "properties": {
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
    },
    "required": [],
    "additionalProperties": False,
}

MD_TO_SPREADSHEET_TEMPLATE_OPTIONS_SCHEMA: dict = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "template_name": {
            "type": "string",
            "pattern": r"^template\.xlsx\.[0-9a-f]{64}$",
            "description": (
                "Canonical XLSX template resource ID published by resources list templates. "
                "Names, filenames, paths, and blank values are rejected. "
                "Supported for XLSX and spreadsheet targets that render through an XLSX intermediate."
            ),
        },
    },
    "required": [],
}

MD_TO_SPREADSHEET_OPTIONS_SCHEMA: dict = {
    "type": "object",
    "additionalProperties": False,
    "properties": {},
    "required": [],
}

MD_NUMBERING_OPTIONS_SCHEMA: dict = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "remove_numbering": {
            "type": "boolean",
            "default": True,
            "description": "Remove existing heading numbering from Markdown.",
        },
        "add_numbering": {
            "type": "boolean",
            "default": False,
            "description": "Add heading numbering to Markdown headings.",
        },
        "numbering_scheme": {
            "type": "string",
            "enum": [
                "gongwen_standard",
                "hierarchical_standard",
                "legal_standard",
            ],
            "default": "gongwen_standard",
            "description": "Numbering scheme ID (e.g. gongwen_standard).",
        },
    },
    "required": [],
}

# Shared empty options schema for routes without user-facing options.
_NOT_IMPL_OPTIONS_SCHEMA: dict = {
    "type": "object",
    "additionalProperties": False,
    "properties": {},
    "required": [],
}


# ── Conversion routes ──────────────────────────────────────────────────

ROUTE_MD_TO_DOCX = RouteSpec(
    source_format="markdown",
    target_format="docx",
    action_name="",
    label="Markdown → DOCX",
    options_schema=MD_TO_DOCX_OPTIONS_SCHEMA,
)

ROUTE_MD_TO_DOC = RouteSpec(
    source_format="markdown",
    target_format="doc",
    action_name="",
    label="Markdown → DOC (Office bridge)",
    options_schema=MD_TO_DOCX_OPTIONS_SCHEMA,
)

ROUTE_MD_TO_ODT = RouteSpec(
    source_format="markdown",
    target_format="odt",
    action_name="",
    label="Markdown → ODT (Office bridge)",
    options_schema=MD_TO_DOCX_OPTIONS_SCHEMA,
)

ROUTE_MD_TO_RTF = RouteSpec(
    source_format="markdown",
    target_format="rtf",
    action_name="",
    label="Markdown → RTF (Office bridge)",
    options_schema=MD_TO_DOCX_OPTIONS_SCHEMA,
)

ROUTE_MD_TO_WPS = RouteSpec(
    source_format="markdown",
    target_format="wps",
    action_name="",
    label="Markdown → WPS (Office bridge)",
    options_schema=MD_TO_DOCX_OPTIONS_SCHEMA,
)

ROUTE_MD_TO_PDF = RouteSpec(
    source_format="markdown",
    target_format="pdf",
    action_name="",
    label="Markdown → PDF (Office bridge)",
    options_schema=MD_TO_DOCX_OPTIONS_SCHEMA,
)

ROUTE_MD_TO_XLSX = RouteSpec(
    source_format="markdown",
    target_format="xlsx",
    action_name="",
    label="Markdown → XLSX",
    options_schema=MD_TO_SPREADSHEET_TEMPLATE_OPTIONS_SCHEMA,
)

ROUTE_MD_TO_XLS = RouteSpec(
    source_format="markdown",
    target_format="xls",
    action_name="",
    label="Markdown → XLS (Office bridge)",
    options_schema=MD_TO_SPREADSHEET_TEMPLATE_OPTIONS_SCHEMA,
)

ROUTE_MD_TO_ODS = RouteSpec(
    source_format="markdown",
    target_format="ods",
    action_name="",
    label="Markdown → ODS (Office bridge)",
    options_schema=MD_TO_SPREADSHEET_TEMPLATE_OPTIONS_SCHEMA,
)

ROUTE_MD_TO_CSV = RouteSpec(
    source_format="markdown",
    target_format="csv",
    action_name="",
    label="Markdown → CSV",
    options_schema=MD_TO_SPREADSHEET_TEMPLATE_OPTIONS_SCHEMA,
)

ROUTE_MD_NUMBERING = RouteSpec(
    source_format="markdown",
    target_format="md",
    action_name="process_md_numbering",
    label="Markdown heading numbering (clean/add)",
    options_schema=MD_NUMBERING_OPTIONS_SCHEMA,
)


# ── All routes ─────────────────────────────────────────────────────────

ALL_ROUTES: list[RouteSpec] = [
    # Document routes
    ROUTE_MD_TO_DOCX,
    ROUTE_MD_TO_DOC,
    ROUTE_MD_TO_ODT,
    ROUTE_MD_TO_RTF,
    ROUTE_MD_TO_WPS,
    ROUTE_MD_TO_PDF,
    # Spreadsheet routes
    ROUTE_MD_TO_XLSX,
    ROUTE_MD_TO_XLS,
    ROUTE_MD_TO_ODS,
    ROUTE_MD_TO_CSV,
    # Action routes
    ROUTE_MD_NUMBERING,
]


def build_manifest() -> PluginManifest:
    """Build and return the Markdown plugin manifest."""
    return PluginManifest(
        plugin_id=PLUGIN_ID,
        name="DocWen Markdown Plugin",
        version=PLUGIN_VERSION,
        description=(
            "Converts Markdown files to DOCX, DOC, ODT, RTF, WPS, PDF, XLSX, XLS, ODS, CSV; "
            "handles heading numbering processing. "
            "Office formats are produced through the core Office bridge."
        ),
        author="DocWen",
        routes=ALL_ROUTES,
        capability_rules=[
            RouteCapabilityRule(required_capabilities=("python.mistune",)),
            RouteCapabilityRule(
                target_formats=("docx", "doc", "odt", "rtf", "wps", "pdf"),
                required_capabilities=("python.docx",),
            ),
            RouteCapabilityRule(
                target_formats=("doc", "odt", "rtf", "wps", "pdf"),
                required_capabilities=("external_office.word",),
                limitations=("the target is published through an external Word-compatible backend",),
            ),
            RouteCapabilityRule(
                target_formats=("xlsx", "xls", "ods", "csv"),
                required_capabilities=("python.openpyxl",),
            ),
            RouteCapabilityRule(
                target_formats=("xls", "ods"),
                required_capabilities=("external_office.spreadsheet",),
                limitations=("the target is published through an external spreadsheet backend",),
            ),
        ],
        extra={
            "third_party_deps": [
                "python-docx (required, for MD→DOCX)",
                "openpyxl (required, for MD→XLSX)",
                "mistune (required, for Markdown parsing)",
            ],
            "office_bridge_routes": [
                HonestyRoute(
                    source="markdown",
                    targets=["doc"],
                    description="via DOCX intermediate + Office bridge (COM/LibreOffice) (ROUTE-MD-DOC-001)",
                ),
                HonestyRoute(
                    source="markdown",
                    targets=["odt"],
                    description="via DOCX intermediate + Office bridge (ROUTE-MD-ODT-001)",
                ),
                HonestyRoute(
                    source="markdown",
                    targets=["rtf"],
                    description="via DOCX intermediate + Office bridge (ROUTE-MD-RTF-001)",
                ),
                HonestyRoute(
                    source="markdown",
                    targets=["wps"],
                    description="via DOCX intermediate + Office bridge (ROUTE-MD-WPS-001)",
                ),
                HonestyRoute(
                    source="markdown",
                    targets=["pdf"],
                    description="via DOCX intermediate + Office bridge (ROUTE-MD-PDF-001)",
                ),
                HonestyRoute(
                    source="markdown",
                    targets=["xls"],
                    description="via XLSX intermediate + Office bridge (ROUTE-MD-XLS-001)",
                ),
                HonestyRoute(
                    source="markdown",
                    targets=["ods"],
                    description="via XLSX intermediate + Office bridge (ROUTE-MD-ODS-001)",
                ),
            ],
        },
    )
