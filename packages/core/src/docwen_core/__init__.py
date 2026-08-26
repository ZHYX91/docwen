"""DocWen Core — shared models, protocols, errors, events, and format definitions.

This package MUST NOT depend on:
- PySide6, python-docx, openpyxl, PyMuPDF, rapidocr, or any specific implementation library
- docwen_application, docwen_runtime, docwen_gui, docwen_cli, docwen_bundle, docwen_plugin_*
"""

from docwen_core.cancellation import CancellationToken
from docwen_core.errors import (
    CancellationRequested,
    ConfigurationError,
    ConversionError,
    DocWenError,
    ValidationError,
)
from docwen_core.links import (
    extract_block_by_id,
    extract_section_by_heading,
    normalize_link_target,
    parse_anchor,
    resolve_file_path,
    strip_yaml_front_matter,
)
from docwen_core.models import (
    PORTABLE_TARGET_ID_MAX_LENGTH,
    ArtifactManifest,
    ConversionDiagnostic,
    ConversionErrorInfo,
    ConversionMetrics,
    ConversionRequest,
    ConversionResult,
    FileRef,
    OptimizationResourceSpec,
    OutputPolicy,
    PluginManifest,
    ProofreadRules,
    RouteSpec,
    SemanticBibliographyEntry,
    SemanticBibliographyFragment,
    SemanticBibliographyRun,
    SemanticCaption,
    SemanticCitationCluster,
    SemanticCitationItem,
    SemanticDiagnostic,
    SemanticDocument,
    SemanticDocumentValidationError,
    SemanticImportResult,
    SemanticParagraph,
    SemanticReference,
    SemanticTable,
    SemanticTableCell,
    SemanticText,
    TaskEvent,
    WorkerRequest,
    WorkerResult,
    derive_table_header_shape,
    is_portable_semantic_id,
    validate_semantic_document,
)
from docwen_core.ofd import apply_easyofd_patches
from docwen_core.protocols import CancellationTokenView
from docwen_core.text import (
    LOCALE_TO_OCR_LANGUAGE,
    OCR_LANGUAGE_AUTO,
    OCR_LANGUAGE_CHINESE,
    OCR_LANGUAGE_CHINESE_CHT,
    OCR_LANGUAGE_CYRILLIC,
    OCR_LANGUAGE_ENGLISH,
    OCR_LANGUAGE_JAPANESE,
    OCR_LANGUAGE_KOREAN,
    OCR_LANGUAGE_LATIN,
    OCR_LANGUAGE_MODELS,
    resolve_ocr_language,
)
from docwen_core.version import PRODUCT_VERSION
from docwen_core.version import __version__ as __version__

__all__ = [  # noqa: RUF022 - grouped by public domain for API readability
    "LOCALE_TO_OCR_LANGUAGE",
    "OCR_LANGUAGE_AUTO",
    "OCR_LANGUAGE_CHINESE",
    "OCR_LANGUAGE_CHINESE_CHT",
    "OCR_LANGUAGE_CYRILLIC",
    "OCR_LANGUAGE_ENGLISH",
    "OCR_LANGUAGE_JAPANESE",
    "OCR_LANGUAGE_KOREAN",
    "OCR_LANGUAGE_LATIN",
    "OCR_LANGUAGE_MODELS",
    # Models
    "ArtifactManifest",
    # Errors
    "CancellationRequested",
    # Cancellation
    "CancellationToken",
    "CancellationTokenView",
    "ConfigurationError",
    "ConversionDiagnostic",
    "ConversionError",
    "ConversionErrorInfo",
    "ConversionMetrics",
    "ConversionRequest",
    "ConversionResult",
    "DocWenError",
    "FileRef",
    "OutputPolicy",
    "OptimizationResourceSpec",
    "PluginManifest",
    "PORTABLE_TARGET_ID_MAX_LENGTH",
    "PRODUCT_VERSION",
    "ProofreadRules",
    "RouteSpec",
    "SemanticBibliographyEntry",
    "SemanticBibliographyFragment",
    "SemanticBibliographyRun",
    "SemanticCaption",
    "SemanticCitationCluster",
    "SemanticCitationItem",
    "SemanticDiagnostic",
    "SemanticDocument",
    "SemanticDocumentValidationError",
    "SemanticImportResult",
    "SemanticParagraph",
    "SemanticReference",
    "SemanticTable",
    "SemanticTableCell",
    "SemanticText",
    "TaskEvent",
    "ValidationError",
    "WorkerRequest",
    "WorkerResult",
    "derive_table_header_shape",
    "is_portable_semantic_id",
    # OFD
    "apply_easyofd_patches",
    # OCR / text
    # Links
    "extract_block_by_id",
    "extract_section_by_heading",
    "normalize_link_target",
    "parse_anchor",
    "resolve_file_path",
    "resolve_ocr_language",
    "strip_yaml_front_matter",
    "validate_semantic_document",
]
