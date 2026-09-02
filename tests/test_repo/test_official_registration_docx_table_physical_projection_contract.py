"""Fail-closed contracts for VIS-2026-07-17-124 physical evidence."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]
GOLDEN = ROOT / "tests" / "fixtures" / "golden"
FIXTURE = GOLDEN / "old_system_docx_official_registration_table_semantics.json"
REPORT_NAME = "official-registration-docx-table-physical-matrix-2026-07-17.md"


def _fixture() -> dict[str, object]:
    return json.loads(FIXTURE.read_text(encoding="utf-8"))


def _addendum() -> dict[str, object]:
    return _fixture()["physical_matrix_addendum"]


def test_source_identity_geometry_and_external_summary_are_pinned() -> None:
    fixture = _fixture()
    source = fixture["source"]
    assert (source["bytes"], source["sha256"]) == (
        73_173,
        "07aadb0f403095b33c7e1e61c037ae4a8194be660c9f4bc9c6bdb521564e6eea",
    )
    assert source["section_count"] == 13
    assert source["table_count"] == 20
    assert source["normalized_table_cell_position_count"] == 1_143
    assert source["grid_span_count"] == 200
    assert source["vertical_merge_count"] == 23

    addendum = _addendum()
    assert addendum["evidence_id"] == "VIS-2026-07-17-124"
    assert addendum["source_identity_exact_to_semantic_fixture"] is True
    geometry = addendum["source_geometry"]
    assert geometry["page_twips"] == {"width": 11_906, "height": 16_838}
    assert geometry["default_usable_content_twips"] == {
        "width": 8_306,
        "height": 13_958,
    }
    assert geometry["section_5_usable_content_twips"] == {
        "width": 9_353,
        "height": 13_958,
    }
    assert geometry["all_sections_have_positive_usable_content_area"] is True

    external = addendum["external_evidence"]
    assert external["inventory_excluding_five_harness_scripts_summary_and_pycache"] == {
        "files": 200,
        "bytes": 34_810_184,
    }
    assert external["evidence_summary"] == {
        "bytes": 30_613,
        "sha256": "a86753e0b9e7cc4a46248991ccc671db0d636c155d25667d6dacc45f3b2b190d",
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
    ) == (10, 10, 10, 150, 3)
    assert execution["all_contact_sheets_inspected"] is True

    for target, triple in addendum["same_target_three_project"].items():
        assert target in {"doc", "rtf", "odt"}
        assert triple["page_counts"] == [15, 15, 15]
        for key in (
            "page_count_equal",
            "render_size_equal",
            "visible_text_equal",
            "pixels_equal",
        ):
            assert triple[key] is True


def test_source_relative_structural_differences_are_localized_without_overclaim() -> None:
    addendum = _addendum()
    projection = addendum["source_relative_projection"]

    assert projection["doc"]["word_story_paragraph_table_section_projection_exact"] is True
    assert projection["doc"]["source_and_output_pages"] == [15, 15]

    rtf = projection["rtf"]
    assert rtf["table_counts"] == [20, 19]
    assert rtf["source_tables_17_and_18_become_output_table_17_without_content_loss"] is True
    assert rtf["fullwidth_square_brackets_become_tortoise_shell_brackets"] == 4
    assert rtf["all_nonempty_table_cell_text_bracket_canonical_exact_in_order"] is True

    odt = projection["odt"]
    assert odt["paragraph_counts"] == [1_076, 1_075]
    assert odt["nonempty_paragraph_counts"] == [307, 307]
    assert odt["compact_nonempty_paragraphs_exact_in_order"] is True
    assert odt["table_13_has_one_extra_empty_cell"] is True
    assert odt["all_nonempty_table_cell_text_exact_in_order"] is True
    assert projection["rtf_and_odt_visible_existing_header_footer_ranges_exact"] is True
    assert projection["rtf_and_odt_inactive_linked_header_range_difference_sections"] == [
        9,
        10,
        11,
        12,
        13,
    ]

    classification = addendum["classification"]
    assert classification == {
        "focused_complex_registration_table_physical_parity_strong": True,
        "current_only_conversion_object_or_physical_regression_found": False,
        "doc_source_object_projection_exact": True,
        "rtf_shared_table_grouping_and_four_glyph_substitutions_are_accepted_source_fidelity": False,
        "odt_shared_empty_object_and_layout_changes_are_accepted_source_fidelity": False,
        "production_change_made": False,
        "broad_doc_docx_rtf_odt_physical_parity_closed": False,
        "overall_parity_closed": False,
    }


def test_physical_addendum_keeps_golden_inventory_and_process_boundary_closed() -> None:
    addendum = _addendum()
    process = addendum["process_boundary"]
    assert process["pre_existing_process"] == {"name": "wps.exe", "pid": 11_388}
    assert process["conversion_added"] == process["conversion_removed"] == []
    assert process["word_render_added"] == process["word_render_removed"] == []
    assert process["termination_command_used"] is False

    golden_files = sorted(GOLDEN.glob("*.json"))
    assert len(golden_files) == 85
    assert FIXTURE in golden_files
    for path in golden_files:
        json.loads(path.read_text(encoding="utf-8"))

    assert not (ROOT / "tests" / "fixtures" / "files" / "202604301410478373010.docx").exists()
