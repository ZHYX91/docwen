"""Layout plugin manifest — route declarations and option schemas.

Route coverage (explicit source formats so runtime RouteRegistry.find()
can match without category fallback):

    pdf/ofd/xps → md                  (pymupdf4llm)
    pdf/ofd/xps → png                 (PyMuPDF rendering)
    pdf/ofd/xps → jpg                 (PyMuPDF rendering)
    pdf/ofd/xps → tif                 (PyMuPDF rendering)
    pdf → docx                        (configured Office/LibreOffice priority, then pdf2docx fallback)
    ofd/xps → docx                    (Office bridge-backed)
    pdf/ofd/xps → doc                 (Office bridge-backed)
    pdf/ofd/xps → odt                 (Office bridge-backed)
    pdf/ofd/xps → rtf                 (Office bridge-backed)
    pdf/ofd/xps → pdf                 (PyMuPDF / normalize)
    merge_pdfs                        (PyMuPDF)
    split_pdf                         (PyMuPDF)
"""

from __future__ import annotations

from docwen_core.models.manifest import HonestyRoute, PluginManifest, RouteCapabilityRule, RouteSpec

PLUGIN_ID = "docwen_plugin_layout"
PLUGIN_VERSION = "0.1.0"

# ── Option schemas (JSON Schema dicts) ────────────────────────────────

LAYOUT_TO_DOCUMENT_OPTIONS_SCHEMA: dict = {
    "type": "object",
    "additionalProperties": False,
    "properties": {},
    "required": [],
}

LAYOUT_TO_IMAGE_OPTIONS_SCHEMA: dict = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "render_dpi": {
            "type": "integer",
            "default": 150,
            "description": "Rendering DPI for page-to-image conversion (72–600)",
            "minimum": 72,
            "maximum": 600,
        },
    },
    "required": [],
}

LAYOUT_TO_MD_OPTIONS_SCHEMA: dict = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "to_md_keep_images": {
            "type": "boolean",
            "default": True,
            "description": "Export extracted images independently of physical-page OCR",
        },
        "to_md_enable_ocr": {
            "type": "boolean",
            "default": False,
            "description": "Emit exactly one typed OCR fragment per physical page",
        },
        "ocr_language": {
            "type": "string",
            "enum": ["auto", "chinese", "chinese_cht", "english", "japanese", "korean", "latin", "cyrillic"],
            "default": "auto",
            "description": "OCR language used for independently rendered physical pages.",
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
        },
        "image_link_style": {
            "type": "string",
            "enum": ["wiki_embed", "wiki_link", "markdown_embed", "markdown_link"],
            "default": "wiki_embed",
            "description": "Markdown image link style for generated references.",
        },
        "render_dpi": {
            "type": "integer",
            "default": 200,
            "description": "Rendering DPI for page-to-image conversion",
        },
    },
    "required": [],
}

PDF_OPERATION_OPTIONS_SCHEMA: dict = {
    "type": "object",
    "additionalProperties": False,
    "properties": {},
    "required": [],
}

PDF_SPLIT_OPTIONS_SCHEMA: dict = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "split_mode": {
            "type": "string",
            "enum": ["custom", "every_page", "odd_even"],
            "default": "custom",
            "description": "Split mode",
        },
        "pages": {
            "type": "array",
            "items": {"type": "integer", "minimum": 1},
            "description": "Page numbers for custom split mode",
        },
    },
    "required": [],
}

# ── Source formats covered by the fixed-layout conversion routes ─────

# All concrete admitted source formats that convert() handles for generic targets.
# We register explicit routes for every one
# so that RouteRegistry.find() — which does exact (source, target) matching
# without category fallback — can resolve them.
_FIXED_LAYOUT_SOURCES: tuple[str, ...] = ("pdf", "ofd", "xps")

# Target families and their option schemas.
_IMAGE_TARGETS: tuple[tuple[str, str, dict], ...] = (
    ("png", "PNG", LAYOUT_TO_IMAGE_OPTIONS_SCHEMA),
    ("jpg", "JPG", LAYOUT_TO_IMAGE_OPTIONS_SCHEMA),
    ("tif", "TIF", LAYOUT_TO_IMAGE_OPTIONS_SCHEMA),
)

_DOCUMENT_TARGETS: tuple[tuple[str, str, dict], ...] = (
    ("docx", "DOCX", LAYOUT_TO_DOCUMENT_OPTIONS_SCHEMA),
    ("doc", "DOC", LAYOUT_TO_DOCUMENT_OPTIONS_SCHEMA),
    ("odt", "ODT", LAYOUT_TO_DOCUMENT_OPTIONS_SCHEMA),
    ("rtf", "RTF", LAYOUT_TO_DOCUMENT_OPTIONS_SCHEMA),
)


def _make_route(
    source: str,
    target: str,
    label_target: str,
    options_schema: dict,
    *,
    action_name: str = "",
) -> RouteSpec:
    """Create a RouteSpec with a standardised label."""
    label = f"{source.upper()} → {label_target}"
    if action_name:
        label += f" ({action_name})"
    return RouteSpec(
        source_format=source,
        target_format=target,
        action_name=action_name,
        label=label,
        options_schema=options_schema,
    )


# ── Build all conversion routes ──────────────────────────────────────


def _build_all_routes() -> list[RouteSpec]:
    routes: list[RouteSpec] = []

    for src in _FIXED_LAYOUT_SOURCES:
        # → Markdown
        routes.append(_make_route(src, "md", "MD", LAYOUT_TO_MD_OPTIONS_SCHEMA))

        # → Images
        for tgt, label, schema in _IMAGE_TARGETS:
            routes.append(_make_route(src, tgt, label, schema))

        # → Documents (PDF→DOCX tries configured external Office/LibreOffice
        # priority before the pdf2docx fallback; other fixed-layout document
        # targets continue through the shared Office bridge).
        for tgt, label, schema in _DOCUMENT_TARGETS:
            routes.append(_make_route(src, tgt, label, schema))

        # → PDF (normalize / passthrough)
        routes.append(_make_route(src, "pdf", "PDF", PDF_OPERATION_OPTIONS_SCHEMA))

    # ── Action routes ─────────────────────────────────────────────────
    routes.append(
        RouteSpec(
            source_format="pdf",
            target_format="pdf",
            action_name="merge_pdfs",
            label="Merge PDFs",
            options_schema=PDF_OPERATION_OPTIONS_SCHEMA,
        )
    )
    routes.append(
        RouteSpec(
            source_format="pdf",
            target_format="pdf",
            action_name="split_pdf",
            label="Split PDF",
            options_schema=PDF_SPLIT_OPTIONS_SCHEMA,
        )
    )

    return routes


ALL_ROUTES: list[RouteSpec] = _build_all_routes()


def build_manifest() -> PluginManifest:
    """Build and return the Layout plugin manifest."""
    return PluginManifest(
        plugin_id=PLUGIN_ID,
        name="DocWen Layout Plugin",
        version=PLUGIN_VERSION,
        description=(
            "Converts fixed-layout documents (PDF/OFD/XPS) to images "
            "(PNG/JPG/TIF), document formats (PDF→DOCX via configured Office/LibreOffice priority "
            "with pdf2docx fallback; "
            "DOC/ODT/RTF and non-PDF DOCX routes require external Office software), Markdown, PDF, "
            "and supports PDF merge/split."
        ),
        author="DocWen",
        routes=ALL_ROUTES,
        capability_rules=[
            RouteCapabilityRule(
                source_formats=("ofd",),
                required_capabilities=("python.easyofd",),
                limitations=("OFD input is normalized to PDF before the selected route runs",),
            ),
            RouteCapabilityRule(
                source_formats=("xps",),
                required_capabilities=("python.fitz",),
                limitations=("XPS input is normalized to PDF before the selected route runs",),
            ),
            RouteCapabilityRule(
                target_formats=("md",),
                required_capabilities=("python.pymupdf4llm",),
                optional_capabilities=("python.rapidocr",),
                limitations=("OCR options require the optional RapidOCR capability",),
            ),
            RouteCapabilityRule(
                target_formats=("png", "jpg", "tif"),
                required_capabilities=("python.fitz",),
            ),
            RouteCapabilityRule(
                target_formats=("docx",),
                required_capabilities=("backend.pdf_to_docx",),
                limitations=("PDF to DOCX uses an Office backend or the pdf2docx fallback",),
            ),
            RouteCapabilityRule(
                target_formats=("doc", "odt", "rtf"),
                required_capabilities=("external_office.word",),
                limitations=("document output requires a compatible external Office backend",),
            ),
            RouteCapabilityRule(
                action_names=("merge_pdfs", "split_pdf"),
                required_capabilities=("python.fitz",),
            ),
        ],
        extra={
            "third_party_deps": [
                "PyMuPDF (required, for layout→image rendering and PDF operations)",
                "Office bridge backend (Microsoft Office/LibreOffice per software.special_conversions.pdf_to_office before PDF→DOCX pdf2docx fallback; WPS/Microsoft Office/LibreOffice per software.default_priority.word_processors for DOC/RTF; Microsoft Office/LibreOffice per software.special_conversions.odt for ODT, with WPS excluded)",
                "pdf2docx (final fallback path for PDF→DOCX)",
                "pymupdf4llm (required, for layout→md)",
                "easyofd (optional, for OFD→PDF)",
            ],
            "office_bridge_routes": [
                HonestyRoute(
                    source="pdf",
                    targets=["docx", "doc", "odt", "rtf"],
                    description="PDF document outputs use configured Office backends; DOCX may fall back to pdf2docx",
                ),
                HonestyRoute(
                    source="ofd",
                    targets=["docx", "doc", "odt", "rtf"],
                    description="OFD is admitted and preprocessed to PDF before the configured Office-backed document path",
                ),
                HonestyRoute(
                    source="xps",
                    targets=["docx", "doc", "odt", "rtf"],
                    description="XPS is admitted and preprocessed to PDF before the configured Office-backed document path",
                ),
            ],
        },
    )
