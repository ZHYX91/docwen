"""Tests for content-based file format detection (docwen_core.detection).

Covers: magic-byte signature matching, ZIP container inspection, text
format sniffing, text/binary classification, and fail-closed unknown content.
"""

from __future__ import annotations

import codecs
import os
import tempfile
import zipfile

import pytest

from docwen_core.detection import SUPPORTED_EXTENSION_FORMATS, detect_content_format
from docwen_core.detection._sniffing import (
    detect_text_format,
    has_known_signature,
    is_text_file,
)
from docwen_core.models import StructureStatus

pytestmark = pytest.mark.contract


def _write_temp_file(suffix: str, content: bytes) -> str:
    """Write *content* to a temp file with the given *suffix*."""
    fd, path = tempfile.mkstemp(suffix=suffix)
    os.write(fd, content)
    os.close(fd)
    return path


def _write_temp_text(suffix: str, content: str) -> str:
    """Write *content* (str) to a temp file with the given *suffix*."""
    return _write_temp_file(suffix, content.encode("utf-8"))


def _ooxml_entries(file_format: str) -> dict[str, str]:
    main_parts = {
        "docx": (
            "word/document.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            "document",
        ),
        "xlsx": (
            "xl/workbook.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            "workbook",
        ),
        "pptx": (
            "ppt/presentation.xml",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
            "presentation",
        ),
    }
    main_part, content_type, root_name = main_parts[file_format]
    return {
        "[Content_Types].xml": (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            f'<Override PartName="/{main_part}" ContentType="{content_type}"/>'
            "</Types>"
        ),
        "_rels/.rels": (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            f'Target="{main_part}"/>'
            "</Relationships>"
        ),
        main_part: f"<{root_name}/>",
    }


def _odf_entries(file_format: str) -> dict[str, str]:
    mimetypes = {
        "odt": "application/vnd.oasis.opendocument.text",
        "ods": "application/vnd.oasis.opendocument.spreadsheet",
        "odp": "application/vnd.oasis.opendocument.presentation",
    }
    return {
        "mimetype": mimetypes[file_format],
        "META-INF/manifest.xml": (
            '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">'
            '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>'
            "</manifest:manifest>"
        ),
        "content.xml": ('<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>'),
    }


def _write_zip_entries(path: str, entries: dict[str, str]) -> None:
    with zipfile.ZipFile(path, "w") as package:
        for name, payload in entries.items():
            package.writestr(name, payload)


__all__ = (
    "SUPPORTED_EXTENSION_FORMATS",
    "StructureStatus",
    "_odf_entries",
    "_ooxml_entries",
    "_write_temp_file",
    "_write_temp_text",
    "_write_zip_entries",
    "codecs",
    "detect_content_format",
    "detect_text_format",
    "has_known_signature",
    "is_text_file",
    "os",
    "pytest",
    "pytestmark",
    "tempfile",
    "zipfile",
)
