"""Content inspection, filename declarations, and admission decisions.

``inspect_file`` is the canonical ingress API. ``detect_content_format`` is
the lower-level content-only detector. ``SUPPORTED_EXTENSION_FORMATS`` and
``has_supported_filename_declaration`` describe filename declarations for filters and
comparison only; neither is evidence that a file may execute.
"""

from __future__ import annotations

from docwen_core.detection._sniffing import (
    SUPPORTED_EXTENSION_FORMATS,
    detect_content_format,
)
from docwen_core.detection._validation import (
    FileAdmissionError,
    FileAdmissionPathError,
    admission_error_type,
    enforce_file_admission,
    has_supported_filename_declaration,
    inspect_file,
)
from docwen_core.detection.ooxml_signature import (
    OOXML_SIGNATURE_DERIVED_OUTPUT_UNSIGNED,
    OOXML_SIGNATURE_INFO_METADATA_KEY,
    OOXML_SIGNATURE_VALIDATION_UNAVAILABLE,
    OoxmlSignatureInfo,
    freeze_ooxml_signature_info,
    inspect_ooxml_signature_graph,
    signature_derived_output_diagnostic,
    signature_info_for_ref,
    signature_validation_diagnostic,
)

__all__ = [
    "OOXML_SIGNATURE_DERIVED_OUTPUT_UNSIGNED",
    "OOXML_SIGNATURE_INFO_METADATA_KEY",
    "OOXML_SIGNATURE_VALIDATION_UNAVAILABLE",
    "SUPPORTED_EXTENSION_FORMATS",
    "FileAdmissionError",
    "FileAdmissionPathError",
    "OoxmlSignatureInfo",
    "admission_error_type",
    "detect_content_format",
    "enforce_file_admission",
    "freeze_ooxml_signature_info",
    "has_supported_filename_declaration",
    "inspect_file",
    "inspect_ooxml_signature_graph",
    "signature_derived_output_diagnostic",
    "signature_info_for_ref",
    "signature_validation_diagnostic",
]
