"""Route manifest for docwen_plugin_image."""

from __future__ import annotations

from docwen_core.models.manifest import PluginManifest, RouteCapabilityRule, RouteSpec

PLUGIN_ID = "docwen_plugin_image"
PLUGIN_VERSION = "0.1.0"

IMAGE_TO_MD_OPTIONS_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "to_md_keep_images": {"type": "boolean", "default": True},
        "to_md_enable_ocr": {
            "type": "boolean",
            "default": True,
            "description": "Run OCR; multi-frame TIFF emits one typed fragment per physical frame",
        },
        "ocr_language": {
            "type": "string",
            "enum": ["auto", "chinese", "chinese_cht", "english", "japanese", "korean", "latin", "cyrillic"],
            "default": "auto",
            "x-docwen-status": "implemented",
            "description": "OCR language used for image recognition.",
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
            "description": "Image storage mode: file=separate file, base64=inline data URI, embed=local placeholder, omit=no image output (HTML comment placeholder only)",
        },
        "ocr_placement": {
            "type": "string",
            "enum": ["image_md", "main_md"],
            "default": "image_md",
            "description": (
                "Single images retain their placement behavior; for multi-frame TIFF both values are "
                "UI preferences over the same typed physical-frame fragments."
            ),
            "x-docwen-status": "implemented",
        },
        "image_link_style": {
            "type": "string",
            "enum": ["wiki_embed", "wiki_link", "markdown_embed", "markdown_link"],
            "default": "wiki_embed",
            "description": "Markdown image link style for generated image references.",
            "x-docwen-status": "implemented",
        },
    },
    "required": [],
}

IMAGE_FORMAT_COMMON_PROPERTIES = {
    "compress_mode": {"type": "string", "enum": ["lossless", "limit_size"], "default": "lossless"},
    "size_limit": {"type": "integer", "default": 200, "minimum": 1},
    "size_unit": {"type": "string", "enum": ["KB", "MB"], "default": "KB"},
}

IMAGE_FORMAT_OPTIONS_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        **IMAGE_FORMAT_COMMON_PROPERTIES,
        "target_format": {
            "type": "string",
            "enum": ["jpg", "jpeg", "png", "gif", "bmp", "tif", "tiff", "webp"],
            "default": "png",
            "description": (
                "Actual output format for the generic image -> image route. "
                "Explicit image -> jpg/png/etc routes use the route target instead."
            ),
        },
    },
    "required": [],
}

EXPLICIT_IMAGE_FORMAT_OPTIONS_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": dict(IMAGE_FORMAT_COMMON_PROPERTIES),
    "required": [],
}

IMAGE_TO_PDF_OPTIONS_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "quality_mode": {
            "type": "string",
            "enum": ["original", "a4", "a3"],
            "default": "original",
            "description": "PDF page size mode.",
        },
    },
    "required": [],
}

IMAGE_MERGE_OPTIONS_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "mode": {"type": "string", "enum": ["smart", "rgb", "RGB"], "default": "smart"},
        "keep_alpha": {"type": "boolean", "default": True},
    },
    "required": [],
}

ROUTE_IMAGE_TO_MD = RouteSpec(
    source_format="image",
    target_format="md",
    action_name="",
    label="Image → Markdown",
    options_schema=IMAGE_TO_MD_OPTIONS_SCHEMA,
)

ROUTE_IMAGE_TO_PDF = RouteSpec(
    source_format="image",
    target_format="pdf",
    action_name="",
    label="Image → PDF",
    options_schema=IMAGE_TO_PDF_OPTIONS_SCHEMA,
)

ROUTE_IMAGE_TO_IMAGE = RouteSpec(
    source_format="image",
    target_format="image",
    action_name="",
    label="Image format conversion",
    options_schema=IMAGE_FORMAT_OPTIONS_SCHEMA,
)

ROUTE_MERGE_IMAGES_TO_TIFF = RouteSpec(
    source_format="image",
    target_format="tif",
    action_name="merge_images_to_tiff",
    label="Merge images to multipage TIFF",
    options_schema=IMAGE_MERGE_OPTIONS_SCHEMA,
)

# Explicit target format routes for route resolver parity and can_handle consistency.
EXPLICIT_IMAGE_FORMAT_ROUTES = [
    RouteSpec("image", target, "", f"Image → {target.upper()}", EXPLICIT_IMAGE_FORMAT_OPTIONS_SCHEMA)
    for target in ("jpg", "jpeg", "png", "gif", "bmp", "tif", "tiff", "webp")
]

ALL_ROUTES = [
    ROUTE_IMAGE_TO_MD,
    ROUTE_IMAGE_TO_PDF,
    ROUTE_IMAGE_TO_IMAGE,
    *EXPLICIT_IMAGE_FORMAT_ROUTES,
    ROUTE_MERGE_IMAGES_TO_TIFF,
]


def build_manifest() -> PluginManifest:
    """Build the image plugin manifest."""
    return PluginManifest(
        plugin_id=PLUGIN_ID,
        name="DocWen Image Plugin",
        version=PLUGIN_VERSION,
        description="Converts images to Markdown/PDF/other image formats and merges images to TIFF",
        author="DocWen",
        routes=ALL_ROUTES,
        capability_rules=[
            RouteCapabilityRule(
                required_capabilities=("python.pillow",),
                optional_capabilities=("python.pillow_heif",),
                limitations=("HEIC and HEIF inputs require the optional pillow-heif capability",),
            ),
            RouteCapabilityRule(
                target_formats=("md",),
                optional_capabilities=("python.rapidocr",),
                limitations=("OCR options require the optional RapidOCR capability",),
            ),
        ],
        extra={
            "third_party_deps": ["Pillow", "pillow-heif", "img2pdf", "rapidocr_onnxruntime"],
        },
    )
