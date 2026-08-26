from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract
ROOT = Path(__file__).resolve().parents[2]
FIXTURE = ROOT / "tests/fixtures/golden/old_system_apache_poi_typed_validation_semantics.json"


def _data() -> dict:
    return json.loads(FIXTURE.read_text(encoding="utf-8"))


def test_poi_sources_are_frozen_screened_and_not_distributed() -> None:
    data = _data()
    repo = data["source_repository"]
    assert (repo["owner"], repo["repository"], repo["commit"]) == (
        "apache",
        "poi",
        "86e967d9b28d6a322a87ae8fcbf2a7eeb56cef96",
    )
    assert (repo["workbook_paths_screened"], repo["unique_workbook_blobs"]) == (383, 378)
    assert repo["features_screened_before_docwen_execution"] is True
    expected = {
        "DataValidationEvaluations.xlsx": (13679, "ea5706d9afc031a8c2afef880559483bad30ab2c18e2857aca5f3cfaa74decf4"),
        "DataValidations-49244.xlsx": (10705, "8066b12c878d74aa35b55b22dc9daba9a56e91332a35cf96c4fc0f6af6bd3bd1"),
        "dataValidationTableRange.xlsx": (116554, "2a41971477de7c0af993a6a8bc3082145230ef04d087170797d167e563da84da"),
    }
    assert {k: (v["bytes"], v["sha256"]) for k, v in data["sources"].items()} == expected
    for name in expected:
        assert not any((ROOT / "tests/fixtures/files").rglob(name))


def test_typed_validation_breadth_is_real_but_not_markdown_preservation() -> None:
    contract = _data()["validation_contract"]
    assert contract["total_typed_rules"] == 74
    assert contract["types_covered"] == ["custom", "date", "decimal", "list", "textLength", "time", "whole"]
    assert len(contract["comparison_operators_covered"]) == 7
    assert contract["rule_preservation_in_markdown"] is False
    assert "not validation execution" in contract["classification"]
    sources = _data()["sources"]
    assert [sources[name]["validation_count"] for name in sources] == [17, 52, 5]
    assert sources["DataValidations-49244.xlsx"]["representative_rules"][0]["error_title"] == "Invalid Country Code"
    assert sources["dataValidationTableRange.xlsx"]["named_list_formulas"] == [
        "states",
        "years",
        "Measures",
        "highlight",
        "highlight_list",
    ]


def test_poi_nine_production_runs_and_value_projections_match() -> None:
    execution = _data()["execution_contract"]
    assert execution["valid_three_project_executions"] == 9
    assert execution["all_executions_successful"] is True
    expected = {
        "DataValidationEvaluations.xlsx": (
            60,
            1,
            59,
            175,
            36,
            "ce91377bf1e62897792bf54a7016cc6f6c368aabc510de6566ec02b065e53823",
        ),
        "DataValidations-49244.xlsx": (
            35,
            1,
            34,
            193,
            53,
            "210a57abd4a2746055361d9c5c5f6c8b6eb333877f4128f5cb54e67ba9e61ff0",
        ),
        "dataValidationTableRange.xlsx": (
            3340,
            2,
            3338,
            6630,
            46,
            "96f679637454129cd20b8c5d447957fb87bcde93fa5be4eaaa64d68d34bc3815",
        ),
    }
    for name, facts in expected.items():
        item = _data()["normalized_projections"][name]
        assert item["all_three_projects_equal"] is True
        assert (
            item["entries"],
            item["headings"],
            item["table_rows"],
            item["cells"],
            item["blank_cells"],
            item["projection_sha256"],
        ) == facts
        assert item["old_pyside6_current_raw_equal"] is True


def test_large_validation_workbook_retains_source_faithful_integer_improvement() -> None:
    diff = _data()["normalized_projections"]["dataValidationTableRange.xlsx"]["tk_only_numeric_spelling"]
    assert diff == {
        "source_cell": "xdropdown!W2",
        "source_value": 1000,
        "tk": "1000.0",
        "old_pyside6_current": "1000",
        "classification": "old PySide6/current source-faithful presentation improvement",
    }


def test_typed_validation_current_runtime_finalizer_is_complete() -> None:
    runtime = _data()["current_runtime_finalizer"]
    assert runtime["success"] is True and runtime["artifact_count"] == 1
    assert runtime["primary_bytes"] == runtime["output_bytes"] == 195420
    assert runtime["primary_sha256"] == "ac467cb1c83af6fa2f48435a416988f93a84a77643865729eb22683d2c8f5ede"
    assert (
        runtime["sheet_count"],
        runtime["row_count"],
        runtime["column_count"],
        runtime["block_count"],
        runtime["image_count"],
    ) == (2, 3164, 25, 15, 0)
    assert runtime["diagnostics"] == ["SHEET2MD-OK", "FINALIZER_DONE"]
    assert runtime["workspace_paths_leaked"] is False
