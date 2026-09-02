from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

PROJECT_ROOT = Path(__file__).resolve().parents[2]
FIXTURE = PROJECT_ROOT / "tests/fixtures/golden/old_system_official_openxml_chart_missing_cache_semantics.json"


def _fixture() -> dict:
    return json.loads(FIXTURE.read_text(encoding="utf-8"))


def test_openxml_sources_are_official_frozen_screened_and_not_distributed() -> None:
    data = _fixture()
    repository = data["source_repository"]
    assert repository["owner"] == "dotnet"
    assert repository["repository"] == "Open-XML-SDK"
    assert repository["commit"] == "00967dc871f06776ae969762c6703d062308a6c9"
    assert repository["features_screened_before_docwen_execution"] is True
    assert data["screening_contract"] == {
        "workbook_paths": 109,
        "unique_workbook_blobs": 96,
        "unique_chart_assets": 20,
        "chart_parts": 49,
        "unique_data_validation_assets": 1,
        "data_validation_elements": 32,
        "unique_conditional_formatting_assets": 15,
        "selected_sources": 2,
    }
    expected = {
        "basicspreadsheet.xlsx": (
            26610,
            "13221d4d2569b6f8822fc17d9a1720a272628e914dc8025465285059e4b97d92",
        ),
        "missingcalcchainpart.xlsx": (
            51622,
            "8157632b6baee46941f615f98507380b3081269e2b2f8f0fecfb509ad8b04067",
        ),
    }
    assert {name: (source["bytes"], source["sha256"]) for name, source in data["sources"].items()} == expected
    fixture_files = PROJECT_ROOT / "tests/fixtures/files"
    for name in expected:
        assert not any(fixture_files.rglob(name))


def test_openxml_feature_inventory_prevents_chart_and_validation_overclaim() -> None:
    sources = _fixture()["sources"]
    chart = sources["basicspreadsheet.xlsx"]
    assert chart["chart_count"] == 1
    assert chart["charts"] == [
        {
            "sheet": "Sheet1",
            "class": "BarChart",
            "series_count": 1,
            "values": "Sheet1!$G$1:$G$3",
            "anchor": "OneCellAnchor",
        }
    ]
    assert [rule["types"][0] for rule in chart["conditional_formatting"]] == [
        "dataBar",
        "colorScale",
        "cellIs",
        "cellIs",
        "top10",
    ]

    missing = sources["missingcalcchainpart.xlsx"]
    assert missing["data_validation_count"] == missing["validation_elements_with_prompt"] == 32
    assert missing["unique_validation_prompts"] == 28
    assert missing["typed_validation_count"] == 0
    assert missing["validation_elements_with_formula"] == 0
    assert "do not prove list/range/numeric" in missing["validation_boundary"]


def test_openxml_nine_production_executions_are_explicit() -> None:
    execution = _fixture()["execution_contract"]
    assert execution["no_image_three_project_executions"] == 6
    assert execution["chart_image_enabled_three_project_executions"] == 3
    assert execution["valid_three_project_executions"] == 9
    assert execution["all_executions_successful"] is True
    assert execution["ocr_enabled"] is False
    assert execution["office_processes_started"] is False


def test_openxml_three_project_value_projections_and_raw_outputs_are_equal() -> None:
    projections = _fixture()["normalized_projections"]
    expected = {
        "basicspreadsheet.xlsx": (
            46,
            3,
            42,
            105,
            6,
            "bbc0bf99f153b64fee2f4ffe98c5d609c08daaab397d3cab06fe313c8cfc1184",
            "32bda1f8ef5f9fdea705931839b9e7e152d640b062d8908831b307d0a7991a98",
        ),
        "missingcalcchainpart.xlsx": (
            79,
            6,
            73,
            639,
            91,
            "f70007bdc5d26aa51005b695e1eb9dd2a42dab0d99ff4c8d2b874adae12644eb",
            "793a79415e1a9c9d64e3cadd30be9cc0bdd9d33b3cdc4936d3ad02ce93d9db0d",
        ),
    }
    for name, facts in expected.items():
        projection = projections[name]
        assert projection["all_three_projects_equal"] is True
        assert (
            projection["entries"],
            projection["headings"],
            projection["table_rows"],
            projection["cells"],
            projection["blank_cells"],
            projection["projection_sha256"],
            projection["old_pyside6_current_raw"]["sha256"],
        ) == facts
        assert projection["old_pyside6_current_raw_equal"] is True


def test_missing_formula_cache_behavior_is_shared_and_not_formula_evaluation() -> None:
    source = _fixture()["sources"]["missingcalcchainpart.xlsx"]
    assert source["calc_chain_part_present"] is False
    assert (source["formula_count"], source["formula_with_cached_value_count"]) == (218, 182)
    assert source["formula_without_cached_value_count"] == 36
    assert source["formula_by_sheet"]["P&L Template"] == {
        "total": 102,
        "cached": 66,
        "missing": 36,
    }
    contract = _fixture()["missing_cache_contract"]
    assert contract["source_formula_cells_without_cached_values"] == 36
    assert contract["all_three_projects_emit_cached_values_or_blank_not_formula_text"] is True
    assert set(contract["formula_text_leaks"].values()) == {0}
    assert "not formula evaluation" in contract["classification"]


def test_chart_image_option_records_shared_non_extraction_boundary() -> None:
    contract = _fixture()["chart_image_option_contract"]
    assert contract["source_chart_count"] == 1
    assert contract["source_embedded_image_count"] == 0
    assert contract["all_three_image_enabled_executions_successful"] is True
    assert set(contract["image_artifact_count_by_project"].values()) == {0}
    assert contract["docwen-ref-pyside6_artifact_kinds"] == ["primary"]
    assert contract["docwen-current_artifact_kinds"] == ["primary"]
    assert contract["markdown_equal_to_respective_no_image_run"] is True
    assert "chart rendering/extraction remains unproven" in contract["classification"]
