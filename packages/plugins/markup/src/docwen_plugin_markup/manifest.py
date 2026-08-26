"""Route manifest for docwen_plugin_markup.

Declared routes:
    ROUTE-HTML-001   html -> md
    ROUTE-MHTML-001  mhtml -> md
    ROUTE-HTM-001    htm -> md
    ROUTE-MHT-001    mht -> md
    ROUTE-ENEX-001   enex -> md
    ROUTE-EPUB-001   epub -> md
"""

from __future__ import annotations

from docwen_core.models.manifest import PluginManifest, RouteCapabilityRule, RouteSpec

PLUGIN_ID = "docwen_plugin_markup"
PLUGIN_VERSION = "0.1.0"

_MARKUP_OPTIONS_SCHEMA: dict = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "to_md_keep_images": {
            "type": "boolean",
            "default": True,
            "description": "Extract and keep images from the source.",
        },
        "to_md_enable_ocr": {
            "type": "boolean",
            "default": False,
            "description": "Run OCR on extracted images when converting to Markdown.",
        },
        "image_mode": {
            "type": "string",
            "enum": ["file", "base64", "embed", "omit"],
            "default": "file",
            "description": "Image storage mode: file=separate file, base64=inline data URI, embed=local reference, omit=no image output.",
        },
        "image_link_style": {
            "type": "string",
            "enum": ["wiki_embed", "wiki_link", "markdown_embed", "markdown_link"],
            "default": "wiki_embed",
            "description": "Markdown image link style used for retained local and remote images.",
        },
        "ocr_language": {
            "type": "string",
            "enum": ["auto", "chinese", "chinese_cht", "english", "japanese", "korean", "latin", "cyrillic"],
            "default": "auto",
            "description": "OCR language used for extracted images.",
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
        "ocr_placement": {
            "type": "string",
            "enum": ["image_md", "main_md"],
            "default": "image_md",
            "description": "Place OCR text in per-image Markdown sidecars or inline in the main Markdown.",
        },
    },
    "required": [],
}

ROUTE_HTML_TO_MD = RouteSpec(
    source_format="html",
    target_format="md",
    action_name="",
    label="HTML → Markdown",
    options_schema=_MARKUP_OPTIONS_SCHEMA,
)
ROUTE_MHTML_TO_MD = RouteSpec(
    source_format="mhtml",
    target_format="md",
    action_name="",
    label="MHTML → Markdown",
    options_schema=_MARKUP_OPTIONS_SCHEMA,
)
ROUTE_HTM_TO_MD = RouteSpec(
    source_format="htm",
    target_format="md",
    action_name="",
    label="HTM → Markdown",
    options_schema=_MARKUP_OPTIONS_SCHEMA,
)
ROUTE_MHT_TO_MD = RouteSpec(
    source_format="mht",
    target_format="md",
    action_name="",
    label="MHT → Markdown",
    options_schema=_MARKUP_OPTIONS_SCHEMA,
)
ROUTE_ENEX_TO_MD = RouteSpec(
    source_format="enex",
    target_format="md",
    action_name="",
    label="ENEX → Markdown",
    options_schema=_MARKUP_OPTIONS_SCHEMA,
)
ROUTE_EPUB_TO_MD = RouteSpec(
    source_format="epub",
    target_format="md",
    action_name="",
    label="EPUB → Markdown",
    options_schema=_MARKUP_OPTIONS_SCHEMA,
)

ALL_ROUTES: list[RouteSpec] = [
    ROUTE_HTML_TO_MD,
    ROUTE_MHTML_TO_MD,
    ROUTE_HTM_TO_MD,
    ROUTE_MHT_TO_MD,
    ROUTE_ENEX_TO_MD,
    ROUTE_EPUB_TO_MD,
]


def build_manifest() -> PluginManifest:
    return PluginManifest(
        plugin_id=PLUGIN_ID,
        name="DocWen Markup Plugin",
        version=PLUGIN_VERSION,
        description=(
            "Converts structured text / container formats to Markdown: "
            "HTML/MHTML/HTM/MHT (web_archive), ENEX (note_export), EPUB (publication)."
        ),
        author="DocWen",
        routes=ALL_ROUTES,
        capability_rules=[
            RouteCapabilityRule(
                required_capabilities=("python.markdownify",),
                optional_capabilities=("python.rapidocr",),
                limitations=("OCR options require the optional RapidOCR capability",),
            ),
            RouteCapabilityRule(
                source_formats=("epub",),
                required_capabilities=("python.ebooklib", "python.bs4"),
            ),
        ],
        extra={
            "third_party_deps": ["markdownify", "ebooklib (optional)", "beautifulsoup4 (optional)"],
        },
    )
