"""Tests for the shared external-office bridge."""

from __future__ import annotations

import os
import tempfile
import zipfile
from pathlib import Path
from urllib import parse as urllib_parse

import pytest

pytestmark = pytest.mark.unit


def _write_formula_writer_container(
    path: Path,
    *,
    document_format: str,
    font_name: str = "Times New Roman",
) -> None:
    with zipfile.ZipFile(path, "w") as archive:
        if document_format == "docx":
            archive.writestr(
                "word/document.xml",
                f"""<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
                    <w:body><w:r><w:rPr><w:rFonts w:ascii="{font_name}"/></w:rPr>
                    <m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></w:r></w:body>
                    </w:document>""",
            )
        else:
            archive.writestr(
                "META-INF/manifest.xml",
                """<manifest:manifest
                    xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
                    <manifest:file-entry manifest:full-path="Object 1/"
                    manifest:media-type="application/vnd.oasis.opendocument.formula"/>
                    </manifest:manifest>""",
            )
            archive.writestr("content.xml", f"<document>{font_name}</document>")
            archive.writestr("styles.xml", f"<styles>{font_name}</styles>")


__all__ = (
    "Path",
    "_write_formula_writer_container",
    "os",
    "pytest",
    "pytestmark",
    "tempfile",
    "urllib_parse",
)
