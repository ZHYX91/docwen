"""Guards for the VIS-102 multilingual Markdown physical matrix."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
FIXTURE_PATH = (
    PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_markdown_multilingual_physical_semantics.json"
)
REPORT_NAME = "markdown-output-multilingual-physical-matrix-2026-07-16.md"


def _read(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def _fixture() -> dict[str, Any]:
    return json.loads(_read(FIXTURE_PATH))


def test_multilingual_physical_fixture_records_complete_three_project_matrix() -> None:
    data = _fixture()

    assert data["golden_id"] == "GOLDEN-001"
    assert data["case_id"] == "old_system_markdown_multilingual_physical_semantics"
    assert data["source"] == {
        "path": "samples/sample.md",
        "size_bytes": 2529,
        "sha256": "8F0E6B330EC6912D8CF694D0B6DDFC3C05601EF73FF0840A9AAA52687E1A8854",
    }
    assert data["projects"]["old_tk"]["head"] == "ec9298286cfe1379d5c5470db381577ea43ca0fa"
    assert data["projects"]["old_pyside6"]["head"] == "63db927c5ded920d4994bfede5c7b34c55e2f43e"
    assert data["projects"]["current"]["head"] == "07e3a02"

    templates = data["template_matrix"]
    assert len(templates) == 11
    assert [item["index"] for item in templates] == list(range(1, 12))
    assert len({item["name"] for item in templates}) == 11
    assert len({item["template_sha256"] for item in templates}) == 11
    assert all(len(item["template_sha256"]) == 64 for item in templates)
    assert all(len(item["current_docx_sha256"]) == 64 for item in templates)

    for item in templates:
        pages = item["pages"]
        assert pages["current_word"] == pages["tk_word"], item["name"]
        assert pages["current_wps"] == pages["current_word"], item["name"]
        assert pages["old_pyside6_word"] == pages["tk_word"] - 2, item["name"]


def test_multilingual_physical_fixture_records_semantic_and_page_integrity() -> None:
    data = _fixture()
    semantic = data["semantic_projection"]
    physical = data["physical_projection"]

    assert semantic["docx_artifact_count"] == 33
    assert semantic["shared_projection_matches"] == 33
    assert semantic["all_sections_a4"] is True
    assert semantic["all_canonical_text_streams_equal"] is True
    assert semantic["all_table_projections_equal"] is True
    assert semantic["all_note_text_projections_equal"] is True
    assert semantic["visible_paragraph_grouping"] == {
        "old_tk": 83,
        "old_pyside6": 83,
        "current": 82,
    }
    assert semantic["old_core_title_boundary"] is True
    assert semantic["current_core_title_complete"] is True

    assert physical["pdf_count"] == 44
    assert physical["rendered_page_count"] == 246
    assert physical["all_pages_a4"] is True
    assert physical["all_pages_nonblank"] is True
    assert physical["no_content_touches_raster_edge"] is True
    assert physical["all_required_text_anchors_present"] is True
    assert physical["no_literal_double_dollar"] is True
    assert physical["current_word_matches_tk_page_count"] == 11
    assert physical["current_word_wps_page_counts_equal"] == 11
    assert physical["current_metadata"] == {
        "title": "Test File",
        "subject": "MD to DOCX Test",
    }
    assert "pypdf found every required anchor in 44/44 PDFs" in physical["text_extractor_boundary"]
    assert "pdfplumber missed" in physical["text_extractor_boundary"]


def test_multilingual_physical_fixture_pins_external_evidence_without_raw_artifacts() -> None:
    evidence = _fixture()["external_evidence"]

    assert evidence["root"] == "D:/docwen-parity/vis102-markdown-multilingual-07e3a02-v1"
    assert evidence["file_count"] == 422
    assert evidence["size_bytes"] == 44620098
    assert evidence["semantic_projection"] == {
        "size_bytes": 879315,
        "sha256": "4AA5CF0C72719709706EA54415BBC9A5B14B13BD036E80B796ABD464C1BA2F58",
    }
    assert evidence["physical_projection"] == {
        "size_bytes": 281757,
        "sha256": "EB5DA1F7C2FC26956555CBB0A9EBF793FA3D763ED51651208B14C43C2E703D12",
    }
    assert evidence["rendered_pngs"] == {
        "count": 270,
        "page_count": 246,
        "size_bytes": 17309275,
        "dpi": 120,
    }
    assert len(evidence["manifests"]) == 7
    assert all(len(item["sha256"]) == 64 for item in evidence["manifests"].values())
