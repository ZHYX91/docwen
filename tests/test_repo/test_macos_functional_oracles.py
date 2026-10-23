"""Reject plausible but incorrect results in the macOS release package checks."""

from __future__ import annotations

import json
from pathlib import Path

import fitz
import pytest
from docx import Document
from PIL import Image
from scripts.release.macos_functional_cases import Case, verify_case

pytestmark = pytest.mark.unit


def test_pdf_oracle_rejects_reversed_pages_even_with_success(tmp_path: Path) -> None:
    path = tmp_path / "result.pdf"
    with fitz.open() as pdf:
        for text in ("SECOND PAGE", "FIRST PAGE"):
            pdf.new_page().insert_text((50, 50), text)
        pdf.save(path)
    with pytest.raises(ValueError, match="wrong merged PDF order"):
        verify_case(Case("merge", (), "merge-pdf", path), {"success": True}, {})


def test_docx_oracle_requires_table_content(tmp_path: Path) -> None:
    path = tmp_path / "result.docx"
    document = Document()
    document.add_paragraph("DOCWEN PROBE CONTENT")
    document.save(path)
    with pytest.raises(ValueError, match="missing DOCX table"):
        verify_case(Case("convert", (), "docx", path), {"success": True}, {})


def test_split_oracle_checks_selected_then_remaining_group(tmp_path: Path) -> None:
    for number, text in ((1, "SECOND PAGE"), (2, "FIRST PAGE")):
        with fitz.open() as pdf:
            pdf.new_page().insert_text((50, 50), text)
            pdf.save(tmp_path / f"two_part{number}.pdf")
    evidence = verify_case(Case("split", (), "split-pdf", tmp_path), {"success": True}, {})
    assert set(evidence["files"]) == {"two_part1.pdf", "two_part2.pdf"}


def test_office_pdf_oracle_accepts_line_wrapping_and_keeps_observed_text(tmp_path: Path) -> None:
    output = tmp_path / "output.pdf"
    with fitz.open() as pdf:
        pdf.new_page().insert_text((40, 40), "DOCWEN PROBE\nCONTENT")
        pdf.save(output)
    result = verify_case(Case("office", (), "office-pdf", output), {"success": True}, {})
    assert result["observations"]["extractedPdfText"] == "DOCWEN PROBE CONTENT"


def test_unsupported_presentation_route_is_separate_from_compatibility_failure(tmp_path: Path) -> None:
    output = tmp_path / "output"
    payload = {
        "success": False,
        "error": {"code": "invalid_input", "message": "No canonical runtime route for pptx -> pdf"},
    }
    result = verify_case(Case("office", (), "unsupported-target", output, 2), payload, {})
    assert result == {"productRouteUnavailable": True}


def test_tiff_oracle_rejects_a_single_frame(tmp_path: Path) -> None:
    path = tmp_path / "result.tif"
    Image.new("RGB", (32, 24), "red").save(path)
    with pytest.raises(ValueError, match="wrong TIFF frames"):
        verify_case(Case("merge", (), "tiff", path), {"success": True}, {})


def test_numbering_oracle_rejects_unchanged_text(tmp_path: Path) -> None:
    source = tmp_path / "source.md"
    result = tmp_path / "numbered.md"
    source.write_text("# Probe title\n\nDOCWEN PROBE CONTENT\nAlpha", encoding="utf-8")
    result.write_bytes(source.read_bytes())
    with pytest.raises(ValueError, match="numbering did not change"):
        verify_case(Case("number", (), "number", result), {"success": True}, {"md": source})


def test_proofread_oracle_rejects_report_for_different_source(tmp_path: Path) -> None:
    source = tmp_path / "source.md"
    result = tmp_path / "report.json"
    source.write_text("（", encoding="utf-8")
    result.write_text(
        json.dumps({"schema": "docwen.proofread_report.v2", "issues": [{}], "source": {"content_sha256": "0" * 64}}),
        encoding="utf-8",
    )
    with pytest.raises(ValueError, match="wrong proofread source"):
        verify_case(Case("validate", (), "proof-md", result), {"success": True}, {"proof-md": source})


def test_existing_output_refusal_must_preserve_bytes(tmp_path: Path) -> None:
    path = tmp_path / "existing.pdf"
    path.write_bytes(b"REPLACED")
    with pytest.raises(ValueError, match="existing output changed"):
        verify_case(Case("conflict", (), "conflict", path, 7), {"error": {"code": "output_exists"}}, {})
