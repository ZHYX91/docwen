from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract
ROOT = Path(__file__).resolve().parents[2]
FIXTURE = ROOT / "tests/fixtures/golden/old_system_apache_poi_typed_validation_semantics.json"


def _data() -> dict:
    return json.loads(FIXTURE.read_text(encoding="utf-8"))


def _addendum() -> dict:
    return _data()["final_artifact_ui_addendum"]


def test_final_artifact_matrix_and_rule_projection_are_exact() -> None:
    addendum = _addendum()
    assert addendum["stage"] == "VIS-2026-07-17-117"
    routes = addendum["route_contract"]
    assert (routes["source_count"], routes["target_count"], routes["project_count"]) == (3, 2, 3)
    assert (routes["successful_conversions"], routes["valid_conversions"]) == (18, 18)
    assert routes["timings_comparable"] is False

    projection = addendum["excel_com_projection"]
    assert projection["source_rule_counts"] == {
        "DataValidationEvaluations.xlsx": 17,
        "DataValidations-49244.xlsx": 52,
        "dataValidationTableRange.xlsx": 5,
    }
    assert projection["source_validated_cell_counts"] == {
        "DataValidationEvaluations.xlsx": 33,
        "DataValidations-49244.xlsx": 53,
        "dataValidationTableRange.xlsx": 6,
    }
    assert (projection["distinct_source_rules"], projection["output_rule_probes"]) == (74, 444)
    assert projection["output_range_probes"] == 444
    assert projection["all_18_outputs_exact_to_source_rule_projection"] is True
    assert projection["all_six_source_target_cross_project_projections_exact"] is True
    assert projection["fields_compared"] == [
        "address",
        "present",
        "type",
        "operator",
        "formula1",
        "formula2",
        "error_title",
        "error_message",
        "show_error",
        "alert_style",
        "ignore_blank",
        "in_cell_dropdown",
    ]


def test_physical_excel_dropdown_and_invalid_input_are_observed() -> None:
    ui = _addendum()["physical_excel_ui"]
    assert ui["xls_projects_observed"] == [
        "docwen-ref-tk",
        "docwen-ref-pyside6",
        "docwen-current",
    ]
    assert ui["current_ods_observed"] is True
    assert ui["dropdown_values"] == ["IN", "US", "UK"]
    assert (ui["input_prompt_title"], ui["input_prompt_message"]) == (
        "Country Code Selection",
        "Choose a Country Code.",
    )
    assert (ui["invalid_error_title"], ui["invalid_error_message"]) == (
        "Invalid Country Code",
        "The specified country code is not a valid option.",
    )
    assert (ui["dropdown_passes"], ui["invalid_input_rejection_passes"]) == (4, 4)
    assert ui["all_copies_closed_without_save"] is True
    assert ui["post_observation_readonly_b2"] == "IN"
    assert ui["screenshots_persisted"] is False


def test_shared_ods_view_and_name_boundaries_are_not_overclaimed() -> None:
    addendum = _addendum()
    view = addendum["saved_view_projection"]
    assert (
        view["source_xlsx"]
        == view["all_three_xls"]
        == {
            "active_cell": "C37",
            "scroll_row": 15,
            "scroll_column": 1,
            "zoom": 100,
        }
    )
    assert view["all_three_ods"] == {
        "active_cell": "A1",
        "scroll_row": 1,
        "scroll_column": 1,
        "zoom": 100,
    }
    assert view["cross_project_exact_by_target"] is True
    assert "not a current-only regression" in view["classification"]
    boundaries = " ".join(addendum["shared_boundaries_not_accepted"])
    assert "County_Ranking" in boundaries
    assert "mojibake" in boundaries
    assert "not a fidelity oracle" in boundaries
    classification = addendum["classification"]
    assert classification == {
        "current_only_functional_or_ui_regression_found": False,
        "production_change_made": False,
        "focused_typed_validation_final_artifact_evidence": "strong",
        "markdown_validation_preservation": False,
        "broad_spreadsheet_final_artifact_parity": False,
        "overall_parity": False,
    }


def test_external_evidence_identity_is_frozen() -> None:
    evidence = _addendum()["external_evidence"]
    assert (evidence["files"], evidence["bytes"]) == (40, 3267367)
    assert evidence["analysis"] == {
        "bytes": 6256,
        "sha256": "81732fd1cb442db5092b9a622aab2a7ada8bcd7763bb023085ff3390846e72f3",
    }
    assert evidence["validation_projection"] == {
        "bytes": 1346424,
        "sha256": "23f86595f1c3dd9c75eab300d3893a7d2f5db5199b08c0c22b39059bedd87fbb",
    }
    assert evidence["view_projection"] == {
        "bytes": 8922,
        "sha256": "2f3003ca6a7eadf06248bd5429d0e31f064dfab217910663a0b799b6fce6400c",
    }
    assert evidence["ui_observations"] == {
        "bytes": 2804,
        "sha256": "dae21c2dbe15a473e79f5f602e46566b80b3f341cb2eeab0dd73404f4ae03aea",
    }
