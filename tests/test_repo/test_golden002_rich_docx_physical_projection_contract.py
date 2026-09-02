"""Fail-closed contracts for VIS-2026-07-17-125 physical evidence."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]
GOLDEN = ROOT / "tests" / "fixtures" / "golden"
FIXTURE = GOLDEN / "old_system_docx_to_markdown_rich_semantics.json"
REPORT_NAME = "golden002-rich-docx-physical-matrix-2026-07-17.md"


def _fixture() -> dict[str, object]:
    return json.loads(FIXTURE.read_text(encoding="utf-8"))


def _addendum() -> dict[str, object]:
    return _fixture()["physical_matrix_addendum"]


def test_source_identity_provenance_geometry_and_external_summary_are_pinned() -> None:
    addendum = _addendum()
    assert addendum["evidence_id"] == "VIS-2026-07-17-125"
    assert addendum["source_identity"] == {
        "bytes": 38_864,
        "sha256": "01cf29adae831f8b8c3e1af7285f83d9f36132750ce4e9af57b11e2376970b5f",
        "exact_to_surviving_transient_probe": True,
        "classification": "controlled generated probe, not an official real-world document",
    }
    assert addendum["source_geometry"] == {
        "sections": 1,
        "page_twips": {"width": 12_240, "height": 15_840},
        "usable_content_twips": {"width": 8_640, "height": 12_960},
        "positive_usable_content_area": True,
    }
    assert addendum["source_package"] == {
        "media_parts": 1,
        "drawingml_drawings": 1,
        "vml_pictures": 1,
        "vml_textboxes": 1,
        "omml_equations": 1,
        "footnote_references": 1,
        "endnote_references": 1,
        "nested_table_elements": 2,
    }
    external = addendum["external_evidence"]
    assert external["inventory_excluding_five_harness_scripts_summary_and_pycache"] == {
        "files": 60,
        "bytes": 3_551_199,
    }
    assert external["evidence_summary"] == {
        "bytes": 27_395,
        "sha256": "adfccd2384460f9a076f8cd2516ea558de935c40527a01e38ab63356e0aad598",
    }
    assert len(external["key_file_sha256"]) == 9
    assert len(external["contact_sheet_sha256"]) == 3


def test_nine_conversions_and_same_target_triples_are_physically_exact() -> None:
    addendum = _addendum()
    execution = addendum["execution"]
    assert execution["targets"] == ["doc", "rtf", "odt"]
    assert execution["production_conversions"] == 9
    assert execution["all_conversions_successful"] is True
    assert execution["all_target_signatures_valid"] is True
    assert (
        execution["word_documents_opened"],
        execution["word_pdf_exports"],
        execution["rendered_views"],
        execution["rendered_pages"],
        execution["contact_sheets"],
    ) == (10, 10, 10, 10, 3)
    assert execution["all_contact_sheets_inspected"] is True

    for target, triple in addendum["same_target_three_project"].items():
        assert target in {"doc", "rtf", "odt"}
        assert triple["page_counts"] == [1, 1, 1]
        for key in (
            "page_count_equal",
            "render_size_equal",
            "visible_text_equal",
            "pixels_equal",
            "complete_word_projection_equal",
        ):
            assert triple[key] is True


def test_rich_objects_are_preserved_and_shared_source_drift_is_not_accepted() -> None:
    addendum = _addendum()
    projection = addendum["source_relative_projection"]
    assert all(projection["preserved_all_targets"].values())

    for target in ("doc", "rtf"):
        item = projection[target]
        assert item["source_and_output_omath_counts"] == [1, 0]
        assert item["source_and_output_inline_shape_counts"] == [1, 2]
        assert item["equation_visibly_preserved_as_picture"] is True
        assert item["equation_pdf_text_remains_searchable"] is False

    odt = projection["odt"]
    assert odt["source_and_output_omath_counts"] == [1, 1]
    assert odt["equation_visibly_and_structurally_preserved"] is True
    assert odt["source_and_output_field_counts"] == [1, 0]
    assert odt["note_marker_first_character_changes"] == [
        "Footnote -> 1ootnote",
        "Endnote -> indnote",
    ]

    assert addendum["classification"] == {
        "focused_controlled_rich_media_object_physical_parity_strong": True,
        "current_only_conversion_object_or_physical_regression_found": False,
        "doc_or_rtf_equation_remains_searchable_as_text": False,
        "odt_note_marker_projection_is_source_exact": False,
        "shared_source_relative_differences_are_accepted_source_fidelity": False,
        "official_real_world_rich_media_physical_parity_proven": False,
        "production_change_made": False,
        "broad_doc_docx_rtf_odt_physical_parity_closed": False,
        "overall_parity_closed": False,
    }


def test_official_source_boundary_is_executable_without_version_substitution() -> None:
    boundary = _addendum()["official_rich_media_source_boundary"]
    assert boundary["candidate"] == "Cambridge Rich Picture conversation templates"
    assert (boundary["expected_bytes"], boundary["expected_sha256"]) == (
        59_369,
        "4a0607a5a5c4eaa7da9e88cabe76c69267a8f954af8810de184b887a3dd03915",
    )
    assert "TLS and browser runtime failures" in boundary["current_environment_result"]
    assert boundary["different_version_substitution_allowed"] is False
    assert boundary["official_real_world_rich_media_physical_acceptance_closed"] is False
