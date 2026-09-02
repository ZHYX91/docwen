from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

PROJECT_ROOT = Path(__file__).resolve().parents[2]
FIXTURE = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_docx_official_registration_table_semantics.json"


def _fixture() -> dict:
    return json.loads(FIXTURE.read_text(encoding="utf-8"))


def test_registration_fixture_records_auditable_external_source() -> None:
    source = _fixture()["source"]
    assert source["publisher"] == "无锡市数据局"
    assert source["disclosure_page"] == "https://bigdata.wuxi.gov.cn/doc/2024/08/21/4267062.shtml"
    assert source["attachment_url"].endswith("202604301410478373010.docx")
    assert source["bytes"] == 73173
    assert source["sha256"] == "07aadb0f403095b33c7e1e61c037ae4a8194be660c9f4bc9c6bdb521564e6eea"
    assert source["redistribution"] == "external_binary_retained_only_in_workspace_bound_temporary_probe"
    assert not (PROJECT_ROOT / "tests" / "fixtures" / "files" / "202604301410478373010.docx").exists()


def test_registration_fixture_locks_complex_table_ooxml_profile() -> None:
    source = _fixture()["source"]
    assert source["section_count"] == 13
    assert source["table_count"] == 20
    assert source["table_row_count"] == 259
    assert source["table_shapes_rows_columns"] == [
        [14, 7],
        [20, 4],
        [1, 2],
        [9, 1],
        [1, 2],
        [10, 5],
        [18, 7],
        [21, 5],
        [5, 5],
        [10, 5],
        [6, 4],
        [14, 4],
        [18, 6],
        [20, 5],
        [11, 6],
        [12, 4],
        [17, 4],
        [1, 2],
        [10, 4],
        [21, 4],
    ]
    assert source["normalized_table_cell_position_count"] == 1143
    assert source["unique_nonempty_table_paragraph_count"] == 230
    assert source["grid_span_count"] == 200
    assert source["vertical_merge_count"] == 23


def test_registration_fixture_records_both_red_green_regressions() -> None:
    data = _fixture()
    phantom, formatting = data["pre_fix_current_regressions"]
    assert phantom["affected_table_indices_one_based"] == [1, 2, 6, 13]
    assert phantom["source_and_old_pyside6_column_counts"] == [7, 4, 5, 6]
    assert phantom["current_pre_fix_column_counts"] == [10, 6, 7, 7]
    assert formatting["red_expected"] == "| <u>Underlined</u><br>Second paragraph |"
    assert formatting["red_actual"] == "| Underlined Second paragraph |"
    regression_text = json.dumps(data["regressions_fixed"], ensure_ascii=False)
    assert "test_vertical_merge_continuation_does_not_add_a_phantom_column" in regression_text
    assert "test_table_cell_paragraph_breaks_and_formatting_survive_finalizer" in regression_text


def test_registration_post_fix_table_projection_and_runtime_contract_match() -> None:
    data = _fixture()
    normalized = data["normalized_contract"]
    projects = data["post_fix_projects"]
    assert normalized["frontmatter_dictionaries_equal_across_all_outputs"] is True
    assert normalized["all_230_unique_source_table_paragraphs_reachable_across_all_outputs"] is True
    assert normalized["current_post_fix_table_shapes_equal_source_and_old_pyside6"] is True
    assert normalized["current_empty_and_old_pyside6_table_position_count"] == 1143
    assert normalized["current_empty_and_old_pyside6_all_normalized_table_positions_equal"] is True
    assert normalized["current_default_vs_empty_difference_is_configured_merge_fill_only"] is True

    current = projects["docwen-current-explicit-empty"]
    assert current["normalized_cell_positions_equal_to_old_pyside6"] == 1143
    assert current["normalized_cell_position_conflicts"] == 0
    assert current["primary_name"] == "company-registration-application-complete.md"
    assert current["primary_in_requested_output_directory"] is True
    assert current["metadata"] == {
        "paragraph_count": 0,
        "heading_count": 1,
        "table_count": 20,
        "image_count": 0,
    }
    assert current["diagnostic_codes"] == ["DOCX2MD-OK", "FINALIZER_DONE"]
    assert projects["docwen-current-default-fill"]["covered_cells_filled_from_merge_anchor"] == 326


def test_registration_evidence_updates_actual_golden_inventory_only() -> None:
    golden_files = sorted((PROJECT_ROOT / "tests" / "fixtures" / "golden").glob("*.json"))
    assert len(golden_files) == 85
    assert FIXTURE in golden_files
    for path in golden_files:
        json.loads(path.read_text(encoding="utf-8"))
