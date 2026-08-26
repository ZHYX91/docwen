"""Canonical file-inspection and filename-declaration contracts."""

from __future__ import annotations

from pathlib import Path

import pytest

import docwen_core
import docwen_core.detection as detection_api
from docwen_core.detection import has_supported_filename_declaration, inspect_file
from docwen_core.models import AdmissionDecision, FileInspection, FormatRelation, StructureStatus

pytestmark = pytest.mark.contract


@pytest.mark.parametrize(
    ("filename", "content", "detected_format", "detected_category"),
    [
        ("source.pdf", b"%PDF-1.4\n", "pdf", "layout"),
        ("source.png", b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR", "png", "image"),
        ("source.jpg", b"\xff\xd8\xff\xe0\x00\x10JFIF\x00", "jpeg", "image"),
        ("source.csv", b"a,b,c\n1,2,3\n4,5,6\n", "csv", "spreadsheet"),
        ("source.md", b"# Heading\n\n**bold** text\n", "markdown", "markdown"),
    ],
)
def test_inspect_file_returns_content_derived_facts(
    tmp_path: Path,
    filename: str,
    content: bytes,
    detected_format: str,
    detected_category: str,
) -> None:
    source = tmp_path / filename
    source.write_bytes(content)

    inspection = inspect_file(str(source))

    assert inspection.detected_format == detected_format
    assert inspection.detected_category == detected_category
    assert inspection.decision is AdmissionDecision.ALLOW
    assert inspection.relation in {FormatRelation.EXACT_MATCH, FormatRelation.EQUIVALENT_ALIAS}


def test_filename_declaration_and_content_fact_remain_separate(tmp_path: Path) -> None:
    source = tmp_path / "not-really-a-document.docx"
    source.write_text("ordinary readable text\n", encoding="utf-8")

    assert has_supported_filename_declaration(str(source)) is True
    inspection = inspect_file(str(source))

    assert inspection.declared_format == "docx"
    assert inspection.declared_category == "document"
    assert inspection.detected_format == "txt"
    assert inspection.workflow_category == "markdown"
    assert inspection.relation is FormatRelation.CROSS_FAMILY_MISMATCH
    assert inspection.decision is AdmissionDecision.REQUIRE_EXPLICIT_ACCEPTANCE


@pytest.mark.parametrize(
    ("filename", "content", "relation", "decision"),
    [
        (
            "layout.txt",
            b"%PDF-1.4\n",
            FormatRelation.CROSS_FAMILY_MISMATCH,
            AdmissionDecision.REQUIRE_EXPLICIT_ACCEPTANCE,
        ),
        (
            "notes.txt",
            b"# Heading\n\ncontent\n",
            FormatRelation.COMPATIBLE_TEXT,
            AdmissionDecision.ALLOW_WITH_WARNING,
        ),
        (
            "photo.jpg",
            b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR",
            FormatRelation.SAME_FAMILY_MISMATCH,
            AdmissionDecision.ALLOW_WITH_WARNING,
        ),
        (
            "layout.unknown",
            b"%PDF-1.4\n",
            FormatRelation.UNRECOGNIZED_EXTENSION,
            AdmissionDecision.REQUIRE_EXPLICIT_ACCEPTANCE,
        ),
    ],
)
def test_inspect_file_expresses_mismatch_policy_without_suffix_fallback(
    tmp_path: Path,
    filename: str,
    content: bytes,
    relation: FormatRelation,
    decision: AdmissionDecision,
) -> None:
    source = tmp_path / filename
    source.write_bytes(content)

    inspection = inspect_file(str(source))

    assert inspection.relation is relation
    assert inspection.decision is decision
    assert (
        inspection.detected_format not in {source.suffix.lstrip("."), "unknown"}
        or relation is FormatRelation.COMPATIBLE_TEXT
    )


@pytest.mark.parametrize(
    ("filename", "content", "reason_code", "structure_status"),
    [
        ("empty.docx", b"", "FILE_EMPTY", StructureStatus.INVALID),
        (
            "unknown.docx",
            b"\x00\x01\x02\x03\x04\x05\x06\x07",
            "FILE_CONTENT_UNRECOGNIZED",
            StructureStatus.UNVERIFIED,
        ),
    ],
)
def test_unverified_content_blocks_even_with_a_supported_suffix(
    tmp_path: Path,
    filename: str,
    content: bytes,
    reason_code: str,
    structure_status: StructureStatus,
) -> None:
    source = tmp_path / filename
    source.write_bytes(content)

    inspection = inspect_file(str(source))

    assert inspection.declared_supported is True
    assert inspection.detected_format == "unknown"
    assert inspection.decision is AdmissionDecision.BLOCK
    assert inspection.reason_code == reason_code
    assert inspection.structure_status is structure_status


def test_file_inspection_serialization_has_only_canonical_names(tmp_path: Path) -> None:
    source = tmp_path / "source.pdf"
    source.write_bytes(b"%PDF-1.4\n")

    data = inspect_file(str(source)).to_dict()

    assert data["declared_format"] == "pdf"
    assert data["detected_format"] == "pdf"
    assert data["declared_category"] == "layout"
    assert data["detected_category"] == "layout"
    assert data["decision"] == "allow"
    assert {
        "actual_format",
        "actual_category",
        "extension_format",
        "extension_category",
        "is_supported",
        "is_valid",
    }.isdisjoint(data)


def test_file_inspection_deserialization_does_not_accept_retired_aliases() -> None:
    inspection = FileInspection.from_dict(
        {
            "file_path": "retired-aliases.bin",
            "extension_format": "docx",
            "extension_category": "document",
            "actual_format": "pdf",
            "actual_category": "layout",
            "is_supported": True,
        }
    )

    assert inspection.declared_format == "unknown"
    assert inspection.declared_category == "other"
    assert inspection.detected_format == "unknown"
    assert inspection.detected_category == "other"
    assert inspection.workflow_category == "other"
    assert inspection.declared_supported is False


def test_core_root_does_not_export_low_level_detection_shortcuts() -> None:
    import docwen_core

    assert not hasattr(docwen_core, "detect_text_format")
    assert not hasattr(docwen_core, "has_known_signature")
    assert not hasattr(docwen_core, "is_text_file")


@pytest.mark.parametrize(
    "filename",
    [
        "source.pdf",
        "source.docx",
        "source.md",
        "source.markdown",
        "source.txt",
        "source.jpg",
        "source.heif",
        "source.xlsx",
        "source.epub",
        "source",
    ],
)
def test_file_picker_declaration_accepts_registered_extensions(filename: str) -> None:
    assert has_supported_filename_declaration(filename) is True


def test_file_picker_declaration_rejects_unknown_extension() -> None:
    assert has_supported_filename_declaration("source.fantasyext") is False


def test_ambiguous_supported_file_shortcut_is_not_public() -> None:
    assert not hasattr(detection_api, "is_supported_file")
    assert not hasattr(docwen_core, "is_supported_file")


@pytest.mark.parametrize(
    "removed_name",
    [
        "detect_actual_file_format",
        "detect_actual_file_format_uncached",
        "validate_file_format",
        "get_actual_file_category",
        "get_file_info",
        "clear_format_cache",
    ],
)
def test_legacy_suffix_fallback_api_is_not_public(removed_name: str) -> None:
    assert not hasattr(detection_api, removed_name)
    assert not hasattr(docwen_core, removed_name)
