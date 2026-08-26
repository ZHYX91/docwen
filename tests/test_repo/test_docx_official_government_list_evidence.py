from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

PROJECT_ROOT = Path(__file__).resolve().parents[2]
FIXTURE = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_docx_official_government_list_semantics.json"


def _fixture() -> dict:
    return json.loads(FIXTURE.read_text(encoding="utf-8"))


def test_official_government_fixture_records_auditable_external_source() -> None:
    source = _fixture()["source"]
    assert source["publisher"] == "金水区人民政府"
    assert source["disclosure_page"].startswith("https://public.jinshui.gov.cn/")
    assert source["attachment_url"].endswith("010941060n2v.docx")
    assert source["bytes"] == 29246
    assert source["sha256"] == "67d955f1ad1c71ca18221ac342093dce90c6d233ae9d48077afbc7ef6ad53a03"
    assert source["redistribution"] == "external_binary_retained_only_in_workspace_bound_temporary_probe"
    assert not (PROJECT_ROOT / "tests" / "fixtures" / "files" / "010941060n2v.docx").exists()


def test_official_government_fixture_locks_ooxml_sentinel_profile() -> None:
    source = _fixture()["source"]
    assert source["paragraph_count"] == 93
    assert source["nonempty_paragraph_count"] == 81
    assert source["normal_style_nonempty_paragraph_count"] == 81
    assert source["outline_level_9_nonempty_paragraph_count"] == 81
    assert source["numpr_paragraph_count"] == 6
    assert source["num_id_1_paragraph_count"] == 1
    assert source["num_id_0_paragraph_count"] == 5
    assert source["numbering_definition"] == {
        "num_id": "1",
        "abstract_num_id": "0",
        "start": 14,
        "num_fmt": "chineseCounting",
        "lvl_text": "（%1）",
    }


def test_official_government_fixture_records_both_red_green_regressions() -> None:
    data = _fixture()
    pre_fix = data["pre_fix_current_regressions"]
    assert pre_fix["paragraph_metadata_count"] == 0
    assert pre_fix["heading_metadata_count"] == 81
    assert pre_fix["markdown_heading_line_count"] == 81
    assert pre_fix["spurious_num_id_zero_bullet_count"] == 5
    assert pre_fix["classification"] == "two_confirmed_current_only_ooxml_sentinel_regressions"

    regression_text = json.dumps(data["regressions_fixed"], ensure_ascii=False)
    for name in (
        "test_extract_outline_level_treats_level_nine_as_body_text",
        "test_word_body_text_outline_sentinel_does_not_become_h6",
        "test_real_outline_level_zero_still_becomes_h1",
        "test_detect_list_item_treats_num_id_zero_as_numbering_disabled",
        "test_body_text_outline_sentinel_stays_body_through_finalizer",
    ):
        assert name in regression_text or name in data["runtime_regression_test"]


def test_official_government_post_fix_bodies_and_runtime_contract_match() -> None:
    data = _fixture()
    normalized = data["normalized_contract"]
    projects = data["post_fix_projects"]
    assert normalized["frontmatter_dictionaries_equal_across_three_projects"] is True
    assert normalized["bodies_equal_after_eof_newline_normalization"] is True
    assert normalized["all_source_nonempty_paragraphs_reachable_across_three_projects"] is True
    assert normalized["source_nonempty_paragraph_count"] == 81
    assert normalized["heading_line_count_per_project"] == 0
    assert normalized["valid_ordered_list_line_count_per_project"] == 1
    assert normalized["spurious_body_bullet_count_per_project"] == 0

    current = projects["docwen-current"]
    assert current["primary_name"] == "zhengzhou-government-document-format-rules.md"
    assert current["primary_in_requested_output_directory"] is True
    assert current["metadata"] == {
        "paragraph_count": 81,
        "heading_count": 0,
        "table_count": 0,
        "image_count": 0,
    }
    assert current["diagnostic_codes"] == ["DOCX2MD-OK", "FINALIZER_DONE"]


def test_official_government_evidence_updates_actual_golden_inventory_only() -> None:
    golden_files = sorted((PROJECT_ROOT / "tests" / "fixtures" / "golden").glob("*.json"))
    assert len(golden_files) == 85
    assert FIXTURE in golden_files
    for path in golden_files:
        json.loads(path.read_text(encoding="utf-8"))
