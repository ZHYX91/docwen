"""Fail-closed contracts for VIS-2026-07-17-123 physical evidence."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]
GOLDEN = ROOT / "tests" / "fixtures" / "golden"
FIXTURE = GOLDEN / "old_system_docx_official_government_list_semantics.json"
REPORT_NAME = "official-government-docx-physical-matrix-2026-07-17.md"


def _fixture() -> dict[str, object]:
    return json.loads(FIXTURE.read_text(encoding="utf-8"))


def _addendum() -> dict[str, object]:
    return _fixture()["real_document_physical_addendum"]


def test_source_identity_geometry_and_external_summary_are_pinned() -> None:
    fixture = _fixture()
    source = fixture["source"]
    assert source["title"] == "郑州市人民政府公文形式与格式细则"
    assert (source["bytes"], source["sha256"]) == (
        29_246,
        "67d955f1ad1c71ca18221ac342093dce90c6d233ae9d48077afbc7ef6ad53a03",
    )

    addendum = _addendum()
    assert addendum["evidence_id"] == "VIS-2026-07-17-123"
    geometry = addendum["source_geometry"]
    assert geometry["page_twips"] == {"width": 11_906, "height": 16_838}
    assert geometry["usable_content_twips"] == {"width": 8_306, "height": 13_958}
    assert geometry["valid_positive_content_area"] is True

    external = addendum["external_evidence"]
    assert external["inventory_excluding_three_harness_scripts_summary_and_pycache"] == {
        "files": 181,
        "bytes": 43_782_412,
    }
    assert external["evidence_summary"] == {
        "bytes": 10_941,
        "sha256": "f7c33cefc7088bd77c802f1fc17f4104ecc616d94279b39169f3b23e7eaade38",
    }
    assert len(external["key_file_sha256"]) == 7
    assert len(external["contact_sheet_sha256"]) == 3


def test_nine_conversions_and_all_same_target_physical_triples_are_exact() -> None:
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
    ) == (10, 10, 10, 133, 3)
    assert execution["all_contact_sheets_inspected"] is True

    triples = addendum["same_target_three_project"]
    assert triples["doc"]["page_counts"] == [13, 13, 13]
    assert triples["rtf"]["page_counts"] == [13, 13, 13]
    assert triples["odt"]["page_counts"] == [14, 14, 14]
    for target in ("doc", "rtf", "odt"):
        for key in (
            "page_count_equal",
            "render_size_equal",
            "visible_text_equal",
            "pixels_equal",
        ):
            assert triples[target][key] is True


def test_source_story_and_unaccepted_shared_fidelity_boundaries_are_explicit() -> None:
    addendum = _addendum()
    objects = addendum["word_object_projection"]
    assert (objects["source_pages"], objects["source_paragraph_count"]) == (13, 93)
    assert objects["main_story_characters"] == 6_464
    assert objects["all_nine_outputs_main_story_exact_to_source"] is True
    assert objects["all_nine_outputs_paragraphs_exact_to_source"] is True
    assert objects["all_nine_outputs_object_counts_exact_to_source"] is True

    physical = addendum["source_relative_physical_projection"]
    assert physical["doc"]["classification"] == "near_source_layout"
    assert physical["doc"]["pdf_joined_text_exact_to_source"] is True
    assert physical["rtf"]["output_pages"] == [13, 13, 13]
    assert "not_accepted_source_fidelity" in physical["rtf"]["classification"]
    assert physical["odt"]["output_pages"] == [14, 14, 14]
    assert "not_accepted_source_fidelity" in physical["odt"]["classification"]

    classification = addendum["classification"]
    assert classification == {
        "focused_official_government_source_physical_parity_strong": True,
        "current_only_conversion_object_or_physical_regression_found": False,
        "shared_rtf_line_reflow_is_accepted_source_fidelity": False,
        "shared_odt_one_page_expansion_is_accepted_source_fidelity": False,
        "production_change_made": False,
        "broad_doc_docx_rtf_odt_physical_parity_closed": False,
        "overall_parity_closed": False,
    }
    process = addendum["process_boundary"]
    assert process["conversion_added_processes"] == []
    assert process["word_render_added_processes"] == []
    assert process["termination_command_used"] is False
