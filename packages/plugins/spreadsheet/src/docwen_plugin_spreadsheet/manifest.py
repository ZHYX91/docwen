"""Route manifest for docwen_plugin_spreadsheet.

Declares all conversion routes this plugin handles.
"""

from __future__ import annotations

from docwen_core.models.manifest import PluginManifest, RouteCapabilityRule, RouteSpec

PLUGIN_ID = "docwen_plugin_spreadsheet"
PLUGIN_VERSION = "0.1.0"

# ── Option schemas ─────────────────────────────────────────────────────

SPREADSHEET_TO_MD_OPTIONS_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "to_md_keep_images": {
            "type": "boolean",
            "default": True,
            "description": "Extract and keep images from the spreadsheet.",
        },
        "to_md_enable_ocr": {
            "type": "boolean",
            "default": False,
            "description": "Run OCR on embedded images and append recognised text as a blockquote. Requires to_md_keep_images=True.",
            "x-docwen-status": "implemented",
        },
        "ocr_language": {
            "type": "string",
            "enum": ["auto", "chinese", "chinese_cht", "english", "japanese", "korean", "latin", "cyrillic"],
            "default": "auto",
            "description": "OCR language used for embedded images.",
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
            "description": "Image storage mode: file=separate files, base64=inline data URI, embed=local placeholder, omit=no image output.",
            "x-docwen-status": "implemented",
        },
        "ocr_placement": {
            "type": "string",
            "enum": ["image_md", "main_md"],
            "default": "image_md",
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
            "description": "How to handle merged cells. fill=repeat anchor value to covered cells, empty=leave covered cells blank, marker=use '<'/ '^' characters for covered cells.",
        },
    },
    "required": [],
}

CSV_TO_XLSX_OPTIONS_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {},
    "required": [],
}

XLSX_TO_CSV_OPTIONS_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {},
    "required": [],
}

TSV_TO_XLSX_OPTIONS_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {},
    "required": [],
}

XLSX_TO_TSV_OPTIONS_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {},
    "required": [],
}

SHEET_FORMAT_CONVERSION_OPTIONS_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {},
    "required": [],
}

XLSX_TO_ODS_OPTIONS_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "spreadsheet_password": {
            "type": "string",
            "format": "password",
            "writeOnly": True,
            "description": "Request-only workbook/sheet protection password; never persisted.",
        },
        "allow_spreadsheet_protection_loss": {
            "type": "boolean",
            "default": False,
            "description": "Explicit consent to publish an ODS without workbook/sheet password protection.",
        },
    },
    "required": [],
}

TABLE_MERGE_OPTIONS_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "merge_mode": {
            "type": "string",
            "enum": ["row", "col", "cell"],
            "default": "cell",
            "description": "Merge mode: row=by row, col=by column, cell=by cell.",
        },
        "offset_range": {
            "type": "integer",
            "default": 10,
            "minimum": 0,
            "maximum": 50,
            "description": "Row/column offset search range for alignment.",
        },
    },
    "required": ["merge_mode"],
}


# ── Route specs ────────────────────────────────────────────────────────

# --- Core conversion routes ---

ROUTE_SHEET_TO_MD = RouteSpec(
    source_format="spreadsheet",
    target_format="md",
    action_name="",
    label="Spreadsheet → Markdown (XLSX, CSV)",
    options_schema=SPREADSHEET_TO_MD_OPTIONS_SCHEMA,
)

ROUTE_XLSX_TO_MD = RouteSpec(
    source_format="xlsx",
    target_format="md",
    action_name="",
    label="XLSX → Markdown (explicit source format)",
    options_schema=SPREADSHEET_TO_MD_OPTIONS_SCHEMA,
)

ROUTE_CSV_TO_MD = RouteSpec(
    source_format="csv",
    target_format="md",
    action_name="",
    label="CSV → Markdown",
    options_schema=SPREADSHEET_TO_MD_OPTIONS_SCHEMA,
)

ROUTE_TSV_TO_MD = RouteSpec(
    source_format="tsv",
    target_format="md",
    action_name="",
    label="TSV → Markdown",
    options_schema=SPREADSHEET_TO_MD_OPTIONS_SCHEMA,
)

# External-office bridge routes for preprocessing to xlsx
ROUTE_XLS_TO_MD = RouteSpec(
    source_format="xls",
    target_format="md",
    action_name="",
    label="XLS → Markdown (via external office bridge)",
    options_schema=SPREADSHEET_TO_MD_OPTIONS_SCHEMA,
)

ROUTE_ODS_TO_MD = RouteSpec(
    source_format="ods",
    target_format="md",
    action_name="",
    label="ODS → Markdown (via external office bridge)",
    options_schema=SPREADSHEET_TO_MD_OPTIONS_SCHEMA,
)

ROUTE_ET_TO_MD = RouteSpec(
    source_format="et",
    target_format="md",
    action_name="",
    label="ET → Markdown (via external office bridge)",
    options_schema=SPREADSHEET_TO_MD_OPTIONS_SCHEMA,
)

ROUTE_CSV_TO_XLSX = RouteSpec(
    source_format="csv",
    target_format="xlsx",
    action_name="",
    label="CSV → XLSX",
    options_schema=CSV_TO_XLSX_OPTIONS_SCHEMA,
)

ROUTE_XLSX_TO_CSV = RouteSpec(
    source_format="xlsx",
    target_format="csv",
    action_name="",
    label="XLSX → CSV",
    options_schema=XLSX_TO_CSV_OPTIONS_SCHEMA,
)

ROUTE_TSV_TO_XLSX = RouteSpec(
    source_format="tsv",
    target_format="xlsx",
    action_name="",
    label="TSV → XLSX",
    options_schema=TSV_TO_XLSX_OPTIONS_SCHEMA,
)

ROUTE_XLSX_TO_TSV = RouteSpec(
    source_format="xlsx",
    target_format="tsv",
    action_name="",
    label="XLSX → TSV",
    options_schema=XLSX_TO_TSV_OPTIONS_SCHEMA,
)

# --- SmartSheetConverter (format interconversion) routes ---
# These routes use xlsx as the hub format.
# Some require external office software (WPS/Excel/LibreOffice)
# and use the shared office bridge.

# Direct (source or target is hub xlsx):
ROUTE_XLSX_TO_XLS = RouteSpec(
    source_format="xlsx",
    target_format="xls",
    action_name="",
    label="XLSX → XLS (SmartConverter, requires external office)",
    options_schema=SHEET_FORMAT_CONVERSION_OPTIONS_SCHEMA,
)

ROUTE_XLSX_TO_ODS = RouteSpec(
    source_format="xlsx",
    target_format="ods",
    action_name="",
    label="XLSX → ODS (SmartConverter, requires external office)",
    options_schema=XLSX_TO_ODS_OPTIONS_SCHEMA,
)

ROUTE_XLSX_TO_ET = RouteSpec(
    source_format="xlsx",
    target_format="et",
    action_name="",
    label="XLSX → ET (SmartConverter, requires external office)",
    options_schema=SHEET_FORMAT_CONVERSION_OPTIONS_SCHEMA,
)

ROUTE_XLS_TO_XLSX = RouteSpec(
    source_format="xls",
    target_format="xlsx",
    action_name="",
    label="XLS → XLSX (SmartConverter, requires external office)",
    options_schema=SHEET_FORMAT_CONVERSION_OPTIONS_SCHEMA,
)

ROUTE_ODS_TO_XLSX = RouteSpec(
    source_format="ods",
    target_format="xlsx",
    action_name="",
    label="ODS → XLSX (SmartConverter, requires external office)",
    options_schema=SHEET_FORMAT_CONVERSION_OPTIONS_SCHEMA,
)

ROUTE_ET_TO_XLSX = RouteSpec(
    source_format="et",
    target_format="xlsx",
    action_name="",
    label="ET → XLSX (SmartConverter, requires external office)",
    options_schema=SHEET_FORMAT_CONVERSION_OPTIONS_SCHEMA,
)

# Two-step (neither source nor target is hub xlsx):
ROUTE_XLS_TO_ODS = RouteSpec(
    source_format="xls",
    target_format="ods",
    action_name="",
    label="XLS → ODS (SmartConverter two-step via xlsx, requires external office)",
    options_schema=SHEET_FORMAT_CONVERSION_OPTIONS_SCHEMA,
)

ROUTE_XLS_TO_ET = RouteSpec(
    source_format="xls",
    target_format="et",
    action_name="",
    label="XLS → ET (SmartConverter two-step via xlsx, requires external office)",
    options_schema=SHEET_FORMAT_CONVERSION_OPTIONS_SCHEMA,
)

ROUTE_ODS_TO_XLS = RouteSpec(
    source_format="ods",
    target_format="xls",
    action_name="",
    label="ODS → XLS (SmartConverter two-step via xlsx, requires external office)",
    options_schema=SHEET_FORMAT_CONVERSION_OPTIONS_SCHEMA,
)

ROUTE_ODS_TO_ET = RouteSpec(
    source_format="ods",
    target_format="et",
    action_name="",
    label="ODS → ET (SmartConverter two-step via xlsx, requires external office)",
    options_schema=SHEET_FORMAT_CONVERSION_OPTIONS_SCHEMA,
)

ROUTE_ET_TO_XLS = RouteSpec(
    source_format="et",
    target_format="xls",
    action_name="",
    label="ET → XLS (SmartConverter two-step via xlsx, requires external office)",
    options_schema=SHEET_FORMAT_CONVERSION_OPTIONS_SCHEMA,
)

ROUTE_ET_TO_ODS = RouteSpec(
    source_format="et",
    target_format="ods",
    action_name="",
    label="ET → ODS (SmartConverter two-step via xlsx, requires external office)",
    options_schema=SHEET_FORMAT_CONVERSION_OPTIONS_SCHEMA,
)

# CSV/TSV outbound (non-xlsx targets, two-step via xlsx):
ROUTE_CSV_TO_XLS = RouteSpec(
    source_format="csv",
    target_format="xls",
    action_name="",
    label="CSV → XLS (SmartConverter two-step via xlsx, requires external office)",
    options_schema=SHEET_FORMAT_CONVERSION_OPTIONS_SCHEMA,
)

ROUTE_CSV_TO_ODS = RouteSpec(
    source_format="csv",
    target_format="ods",
    action_name="",
    label="CSV → ODS (SmartConverter two-step via xlsx, requires external office)",
    options_schema=SHEET_FORMAT_CONVERSION_OPTIONS_SCHEMA,
)

# Non-xlsx inbound → CSV:
ROUTE_XLS_TO_CSV = RouteSpec(
    source_format="xls",
    target_format="csv",
    action_name="",
    label="XLS → CSV (SmartConverter two-step via xlsx, requires external office)",
    options_schema=SHEET_FORMAT_CONVERSION_OPTIONS_SCHEMA,
)

ROUTE_ODS_TO_CSV = RouteSpec(
    source_format="ods",
    target_format="csv",
    action_name="",
    label="ODS → CSV (SmartConverter two-step via xlsx, requires external office)",
    options_schema=SHEET_FORMAT_CONVERSION_OPTIONS_SCHEMA,
)

ROUTE_ET_TO_CSV = RouteSpec(
    source_format="et",
    target_format="csv",
    action_name="",
    label="ET → CSV (SmartConverter two-step via xlsx, requires external office)",
    options_schema=SHEET_FORMAT_CONVERSION_OPTIONS_SCHEMA,
)

# --- Action routes ---

ROUTE_MERGE_TABLES = RouteSpec(
    source_format="spreadsheet",
    target_format="xlsx",
    action_name="merge_tables",
    label="Merge Tables (row/column/cell mode)",
    options_schema=TABLE_MERGE_OPTIONS_SCHEMA,
)


# ── All routes ─────────────────────────────────────────────────────────

ALL_ROUTES = [
    # Core conversion
    ROUTE_SHEET_TO_MD,
    ROUTE_XLSX_TO_MD,
    ROUTE_CSV_TO_MD,
    ROUTE_TSV_TO_MD,
    ROUTE_XLS_TO_MD,
    ROUTE_ODS_TO_MD,
    ROUTE_ET_TO_MD,
    ROUTE_CSV_TO_XLSX,
    ROUTE_XLSX_TO_CSV,
    ROUTE_TSV_TO_XLSX,
    ROUTE_XLSX_TO_TSV,
    # SmartSheetConverter direct (hub=xlsx)
    ROUTE_XLSX_TO_XLS,
    ROUTE_XLSX_TO_ODS,
    ROUTE_XLSX_TO_ET,
    ROUTE_XLS_TO_XLSX,
    ROUTE_ODS_TO_XLSX,
    ROUTE_ET_TO_XLSX,
    # SmartSheetConverter two-step
    ROUTE_XLS_TO_ODS,
    ROUTE_XLS_TO_ET,
    ROUTE_ODS_TO_XLS,
    ROUTE_ODS_TO_ET,
    ROUTE_ET_TO_XLS,
    ROUTE_ET_TO_ODS,
    # CSV/TSV outbound
    ROUTE_CSV_TO_XLS,
    ROUTE_CSV_TO_ODS,
    # Non-xlsx → CSV
    ROUTE_XLS_TO_CSV,
    ROUTE_ODS_TO_CSV,
    ROUTE_ET_TO_CSV,
    # Actions
    ROUTE_MERGE_TABLES,
]


def build_manifest() -> PluginManifest:
    """Build the plugin manifest for docwen_plugin_spreadsheet."""
    return PluginManifest(
        plugin_id=PLUGIN_ID,
        name="DocWen Spreadsheet Plugin",
        version=PLUGIN_VERSION,
        description=(
            "Converts spreadsheets (XLSX, CSV, TSV) to/from Markdown, CSV/TSV ↔ XLSX interconversion, and table merging"
        ),
        author="DocWen",
        routes=ALL_ROUTES,
        capability_rules=[
            RouteCapabilityRule(required_capabilities=("python.openpyxl",)),
            RouteCapabilityRule(
                source_formats=("xls", "ods", "et"),
                required_capabilities=("external_office.spreadsheet",),
                limitations=("the source is normalized through an external spreadsheet backend",),
            ),
            RouteCapabilityRule(
                target_formats=("xls", "ods", "et"),
                required_capabilities=("external_office.spreadsheet",),
                limitations=("the target is published through an external spreadsheet backend",),
            ),
            RouteCapabilityRule(
                target_formats=("md",),
                optional_capabilities=("python.rapidocr",),
                limitations=("OCR options require the optional RapidOCR capability",),
            ),
        ],
        extra={
            "external_bridge_routes": [
                "xls/ods/et -> md",
                "spreadsheet format interconversion via xlsx hub",
            ],
            "third_party_deps": [
                "openpyxl",
                "pandas",
                "tabulate",
            ],
        },
    )
