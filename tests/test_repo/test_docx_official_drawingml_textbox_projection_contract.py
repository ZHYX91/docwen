from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

PROJECT_ROOT = Path(__file__).resolve().parents[2]
FIXTURE = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_docx_official_drawingml_textbox_semantics.json"


def _fixture() -> dict:
    return json.loads(FIXTURE.read_text(encoding="utf-8"))


def test_drawingml_fixture_records_auditable_external_source() -> None:
    source = _fixture()["source"]
    assert source["publisher"] == "Cambridge City Council"
    assert source["disclosure_page"] == "https://www.cambridge.gov.uk/cambridge-rich-picture"
    assert source["attachment_url"].endswith("/rich-picture-templates.docx")
    assert source["bytes"] == 59369
    assert source["sha256"] == "4a0607a5a5c4eaa7da9e88cabe76c69267a8f954af8810de184b887a3dd03915"
    assert source["redistribution"] == "external_binary_retained_only_in_workspace_bound_temporary_probe"
    assert not (PROJECT_ROOT / "tests" / "fixtures" / "files" / "rich-picture-templates.docx").exists()


def test_drawingml_fixture_locks_shape_and_textbox_ooxml_profile() -> None:
    source = _fixture()["source"]
    assert source["top_level_nonempty_paragraph_count"] == 0
    assert source["table_count"] == 0
    assert source["drawing_count"] == 15
    assert source["alternate_content_count"] == 15
    assert source["drawingml_textbox_count"] == 16
    assert source["vml_fallback_textbox_count"] == 16
    assert source["textbox_content_count"] == 32
    assert source["unique_choice_paragraph_count"] == 21
    assert source["media_file_count"] == 0


def test_drawingml_fixture_records_all_three_red_green_regressions() -> None:
    data = _fixture()
    wrapper_ids, manual_breaks, metadata = data["pre_fix_current_regressions"]
    assert wrapper_ids["source_unique_choice_paragraphs"] == 21
    assert wrapper_ids["current_reachable_unique_source_paragraphs"] == 4
    assert manual_breaks["joined_examples"] == [
        "Describe the outcome:- specific",
        "What needs to be in place?(Enablers)",
    ]
    assert metadata == {
        "description": "Exported textbox paragraphs did not contribute to ArtifactManifest paragraph_count.",
        "pre_fix_paragraph_metadata_count": 0,
        "post_fix_paragraph_metadata_count": 21,
    }
    regression_text = json.dumps(data["regressions_fixed"], ensure_ascii=False)
    for name in (
        "test_distinct_drawingml_textboxes_are_not_lost_to_wrapper_id_reuse",
        "test_drawingml_textbox_preserves_manual_line_breaks",
        "test_drawingml_textbox_paragraph_is_counted_through_finalizer",
    ):
        assert name in regression_text


def test_drawingml_post_fix_projection_and_runtime_contract_are_explicit() -> None:
    data = _fixture()
    projects = data["post_fix_projects"]
    normalized = data["normalized_contract"]
    assert normalized["frontmatter_dictionaries_equal_across_three_projects"] is True
    assert normalized["old_outputs_byte_identical"] is True
    assert normalized["current_all_unique_choice_paragraphs_reachable"] is True
    assert normalized["current_preserves_source_unique_paragraph_order"] is True
    assert projects["docwen-ref-tk"]["reachable_unique_source_paragraphs"] == 19
    assert projects["docwen-ref-pyside6"]["missing_unique_source_paragraphs"] == [
        "What needs to be true?",
        "Coordinator",
    ]
    current = projects["docwen-current"]
    assert current["reachable_unique_source_paragraphs"] == 21
    assert current["missing_unique_source_paragraphs"] == []
    assert current["primary_name"] == "cambridge-rich-picture-templates.md"
    assert current["primary_in_requested_output_directory"] is True
    assert current["metadata"] == {
        "paragraph_count": 21,
        "heading_count": 0,
        "table_count": 0,
        "image_count": 0,
    }
    assert current["diagnostic_codes"] == ["DOCX2MD-OK", "FINALIZER_DONE"]


def test_drawingml_evidence_updates_actual_golden_inventory_only() -> None:
    golden_files = sorted((PROJECT_ROOT / "tests" / "fixtures" / "golden").glob("*.json"))
    assert len(golden_files) == 85
    assert FIXTURE in golden_files
    for path in golden_files:
        json.loads(path.read_text(encoding="utf-8"))
