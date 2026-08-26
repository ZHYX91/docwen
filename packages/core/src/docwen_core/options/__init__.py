"""Shared conversion option schemas.

These JSON Schema dictionaries describe common app/config option surfaces.
Route-specific plugin manifests may expose narrower or more detailed schemas
for their own public routes.
"""

from __future__ import annotations

from typing import Any

from docwen_core.text.heading_merge import DEFAULT_HEADING_MERGE_PUNCTUATION

# ── Common options (shared across all converters) ────────────────────

COMMON_OPTIONS_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "to_md_keep_images": {
            "type": "boolean",
            "default": True,
            "description": "Whether to keep image references in markdown output",
        },
        "to_md_enable_ocr": {
            "type": "boolean",
            "default": False,
            "description": "Whether to run OCR on images",
        },
        "image_mode": {
            "type": "string",
            "enum": ["file", "base64", "embed", "omit"],
            "default": "file",
            "description": (
                "Image storage mode for Markdown output: file=separate files, "
                "base64=inline data URI, embed=placeholder, omit=no image output"
            ),
        },
        "ocr_placement": {
            "type": "string",
            "enum": ["image_md", "main_md"],
            "default": "main_md",
            "description": "Where to place OCR text relative to images",
        },
        "ocr_language": {
            "type": "string",
            "enum": ["auto", "chinese", "chinese_cht", "english", "japanese", "korean", "latin", "cyrillic"],
            "default": "auto",
            "description": "OCR language used for extracted images",
        },
        "image_link_style": {
            "type": "string",
            "enum": ["wiki_embed", "wiki_link", "markdown_embed", "markdown_link"],
            "default": "wiki_embed",
            "description": "Markdown image link style for generated image references",
        },
        "table_merge_strategy": {
            "type": "string",
            "enum": ["fill", "empty", "marker"],
            "default": "fill",
            "description": "How merged table cells are rendered in Markdown tables",
        },
    },
}

# ── Markdown options ─────────────────────────────────────────────────

MARKDOWN_OPTIONS_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "remove_numbering": {
            "type": "boolean",
            "default": True,
            "description": "Remove existing numbering from headings",
        },
        "add_numbering": {
            "type": "boolean",
            "default": False,
            "description": "Add hierarchical numbering to headings",
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
            "description": "Default numbering scheme",
        },
        "formatting_mode": {
            "type": "string",
            "enum": ["apply", "keep", "remove"],
            "default": "apply",
            "description": "How to handle Markdown inline formatting when writing DOCX",
        },
        "heading_formatting_mode": {
            "type": "string",
            "enum": ["apply", "keep", "remove"],
            "default": "remove",
            "description": "How to handle inline formatting inside Markdown headings",
        },
        "table_header_formatting_mode": {
            "type": "string",
            "enum": ["apply", "keep", "remove"],
            "default": "remove",
            "description": "How to handle inline formatting inside Markdown table headers",
        },
        "heading_merge_mode": {
            "type": "string",
            "enum": ["punct_required", "always", "never"],
            "default": "punct_required",
            "description": "When to merge adjacent headings",
        },
        "heading_merge_punctuation": {
            "type": "string",
            "default": DEFAULT_HEADING_MERGE_PUNCTUATION,
            "description": "Editable merge-trigger characters for punct_required mode",
        },
        "list_separator": {
            "type": "string",
            "default": "、",
            "description": "Separator for YAML list values in template placeholders",
        },
    },
}

# ── DOCX options ─────────────────────────────────────────────────────

DOCX_OPTIONS_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "to_md_keep_images": {
            "type": "boolean",
            "default": True,
        },
        "to_md_enable_ocr": {
            "type": "boolean",
            "default": False,
        },
        "image_mode": {
            "type": "string",
            "enum": ["file", "base64", "embed", "omit"],
            "default": "file",
            "description": (
                "Image storage mode for Markdown output: file=separate files, "
                "base64=inline data URI, embed=placeholder, omit=no image output"
            ),
        },
        "ocr_placement": {
            "type": "string",
            "enum": ["image_md", "main_md"],
            "default": "main_md",
            "description": "Where to place OCR text relative to images",
        },
        "ocr_language": {
            "type": "string",
            "enum": ["auto", "chinese", "chinese_cht", "english", "japanese", "korean", "latin", "cyrillic"],
            "default": "auto",
            "description": "OCR language used for extracted images",
        },
        "image_link_style": {
            "type": "string",
            "enum": ["wiki_embed", "wiki_link", "markdown_embed", "markdown_link"],
            "default": "wiki_embed",
            "description": "Markdown image link style for generated image references",
        },
        "table_merge_strategy": {
            "type": "string",
            "enum": ["fill", "empty", "marker"],
            "default": "fill",
            "description": "How merged table cells are rendered in Markdown tables",
        },
        "table_merge": {
            "type": "string",
            "enum": ["empty", "fill"],
            "default": "empty",
            "description": "Strategy for merged table cells",
        },
        "remove_numbering": {
            "type": "boolean",
            "default": True,
            "description": "Remove Word auto-numbering",
        },
    },
}

# ── PDF options ──────────────────────────────────────────────────────

PDF_OPTIONS_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "render_dpi": {
            "type": "integer",
            "default": 300,
            "description": "DPI for rendering PDF pages to images",
        },
        "enable_ocr": {
            "type": "boolean",
            "default": False,
        },
    },
}

# ── Image options ────────────────────────────────────────────────────

IMAGE_OPTIONS_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "compress_mode": {
            "type": "string",
            "enum": ["lossless", "lossy", "none"],
            "default": "lossless",
        },
        "size_limit": {
            "type": "integer",
            "default": 200,
            "description": "Target size limit in KB for lossy compression",
        },
        "ocr_language": {
            "type": "string",
            "default": "auto",
        },
        "tiff_mode": {
            "type": "string",
            "enum": ["smart", "rgb", "keep"],
            "default": "smart",
        },
        "pdf_quality": {
            "type": "string",
            "enum": ["original", "high", "normal"],
            "default": "original",
        },
    },
}

# ── Proofread options ────────────────────────────────────────────────

PROOFREAD_OPTIONS_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "enable_symbol_pairing": {"type": "boolean", "default": True},
        "enable_symbol_correction": {"type": "boolean", "default": True},
        "enable_typos_rule": {"type": "boolean", "default": True},
        "enable_sensitive_word": {"type": "boolean", "default": True},
        "skip_code_blocks": {"type": "boolean", "default": True},
        "skip_quote_blocks": {"type": "boolean", "default": False},
    },
}
