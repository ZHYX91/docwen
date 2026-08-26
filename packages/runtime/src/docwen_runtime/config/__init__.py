"""Configuration registry, loader, and request-policy projections."""

from docwen_runtime.toml_io import atomic_write_text

from .document_styles import DocumentStyleCatalogError, build_document_style_catalog
from .heading_cleanup import build_heading_cleanup_rules
from .loader import RESET_EXCLUDED, ConfigLoader, DocWenConfig
from .ocr_output import build_ocr_blockquote_title
from .proofread_rules import build_proofread_rules

__all__ = [
    "RESET_EXCLUDED",
    "ConfigLoader",
    "DocWenConfig",
    "DocumentStyleCatalogError",
    "atomic_write_text",
    "build_document_style_catalog",
    "build_heading_cleanup_rules",
    "build_ocr_blockquote_title",
    "build_proofread_rules",
]
