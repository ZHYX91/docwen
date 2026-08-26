"""Route manifest for docwen_plugin_presentation.

Declared routes:
    ROUTE-PPTX-001  pptx -> md  (python-pptx)
    ROUTE-PPT-001   ppt -> md   (via external office bridge)
"""

from __future__ import annotations

from docwen_core.models.manifest import PluginManifest, RouteCapabilityRule, RouteSpec

PLUGIN_ID = "docwen_plugin_presentation"
PLUGIN_VERSION = "0.1.0"

PPTX_TO_MD_OPTIONS_SCHEMA: dict = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "export_notes": {"type": "boolean", "default": False},
        "to_md_keep_images": {"type": "boolean", "default": True},
        "to_md_enable_ocr": {"type": "boolean", "default": False},
        "image_mode": {
            "type": "string",
            "enum": ["file", "base64", "embed", "omit"],
            "default": "file",
            "description": "Image storage mode: file=separate file, base64=inline data URI, embed=local reference, omit=no image output.",
        },
        "ocr_placement": {
            "type": "string",
            "enum": ["image_md", "main_md"],
            "default": "main_md",
            "description": "Where OCR text from slide images is written.",
        },
        "ocr_language": {
            "type": "string",
            "enum": ["auto", "chinese", "chinese_cht", "english", "japanese", "korean", "latin", "cyrillic"],
            "default": "auto",
            "description": "OCR language used for slide images.",
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
        "image_link_style": {
            "type": "string",
            "enum": ["wiki_embed", "wiki_link", "markdown_embed", "markdown_link"],
            "default": "wiki_embed",
        },
    },
    "required": [],
}

ROUTE_PPTX_TO_MD = RouteSpec(
    source_format="pptx",
    target_format="md",
    action_name="",
    label="PPTX → Markdown",
    options_schema=PPTX_TO_MD_OPTIONS_SCHEMA,
)

ROUTE_PPT_TO_MD = RouteSpec(
    source_format="ppt",
    target_format="md",
    action_name="",
    label="PPT → Markdown (via external office bridge)",
    options_schema=PPTX_TO_MD_OPTIONS_SCHEMA,
)

ALL_ROUTES: list[RouteSpec] = [
    ROUTE_PPTX_TO_MD,
    ROUTE_PPT_TO_MD,
]


def build_manifest() -> PluginManifest:
    """Build the plugin manifest for docwen_plugin_presentation."""
    return PluginManifest(
        plugin_id=PLUGIN_ID,
        name="DocWen Presentation Plugin",
        version=PLUGIN_VERSION,
        description=("Converts PPTX/PPT presentations to Markdown. PPT→MD requires external office bridge."),
        author="DocWen",
        routes=ALL_ROUTES,
        capability_rules=[
            RouteCapabilityRule(
                required_capabilities=("python.pptx",),
                optional_capabilities=("python.rapidocr",),
                limitations=("OCR options require the optional RapidOCR capability",),
            ),
            RouteCapabilityRule(
                source_formats=("ppt",),
                required_capabilities=("external_office.presentation",),
                limitations=("legacy PPT input is pre-converted by an external presentation backend",),
            ),
        ],
        extra={
            "third_party_deps": ["python-pptx"],
            "external_bridge_routes": ["ppt → md"],
        },
    )
