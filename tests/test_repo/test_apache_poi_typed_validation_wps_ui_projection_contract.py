from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract
ROOT = Path(__file__).resolve().parents[2]
FIXTURE = ROOT / "tests/fixtures/golden/old_system_apache_poi_typed_validation_semantics.json"


def _addendum() -> dict:
    data = json.loads(FIXTURE.read_text(encoding="utf-8"))
    return data["wps_physical_ui_addendum"]


def test_wps_xls_dropdown_and_invalid_input_are_exact_across_projects() -> None:
    addendum = _addendum()
    assert addendum["stage"] == "VIS-2026-07-17-118"
    assert addendum["source_final_artifact_stage"] == "VIS-2026-07-17-117"
    assert addendum["environment"] == {
        "os": "Microsoft Windows 11 Pro 10.0.26200 build 26200",
        "application": "WPS Office Spreadsheet",
        "file_version": "12.1.0.26899",
        "test_owned_copies": True,
        "screenshots_persisted": False,
    }

    ui = addendum["physical_wps_xls"]
    assert ui["b2_projects_observed"] == [
        "docwen-ref-tk",
        "docwen-ref-pyside6",
        "docwen-current",
    ]
    assert ui["dropdown_values"] == ["IN", "US", "UK"]
    assert (ui["input_prompt_title"], ui["input_prompt_message"]) == (
        "Country Code Selection",
        "Choose a Country Code.",
    )
    assert (ui["invalid_error_title"], ui["invalid_error_message"]) == (
        "Invalid Country Code",
        "The specified country code is not a valid option.",
    )
    assert (ui["dropdown_passes"], ui["invalid_input_rejection_passes"]) == (3, 3)
    assert ui["cross_project_behavior_exact"] is True
    assert ui["all_test_copies_closed_without_save"] is True


def test_current_wps_xls_additional_integer_rule_is_not_overclaimed() -> None:
    current = _addendum()["physical_wps_xls"]["current_additional_rule"]
    assert current == {
        "cell": "B6",
        "rule": "whole number greater than zero",
        "input_prompt_title": "Integer Input",
        "input_prompt_message": "Enter an integer value.",
        "invalid_keyboard_value": "0",
        "invalid_error_title": "错误提示",
        "invalid_error_message": "您输入的内容，不符合限制条件。",
        "invalid_value_rejected": True,
        "cross_project_physical_comparison": False,
    }


def test_wps_ods_online_conversion_boundary_is_fail_closed() -> None:
    addendum = _addendum()
    ods = addendum["wps_ods_boundary"]
    assert ods["project_attempted"] == "docwen-current"
    assert ods["opened_locally"] is False
    assert ods["online_service_required"] is True
    assert ods["action"] == "cancelled before allowing conversion"
    assert ods["user_data_transfer_authorized"] is False
    assert ods["cross_project_ods_ui_observed"] is False
    assert "not a current-only regression" in ods["classification"]
    assert "not an ODS UI pass" in ods["classification"]

    classification = addendum["classification"]
    assert classification == {
        "current_only_xls_validation_regression_found": False,
        "focused_wps_xls_physical_parity": "strong",
        "wps_ods_physical_parity": "unavailable_without_online_transfer",
        "production_change_made": False,
        "broad_spreadsheet_final_artifact_parity": False,
        "overall_parity": False,
    }


def test_wps_external_evidence_and_process_boundary_are_frozen() -> None:
    addendum = _addendum()
    assert addendum["external_evidence"] == {
        "root": "D:/docwen-parity/vis118-wps-validation-ui-5fcae33-v1",
        "files": 1,
        "bytes": 4752,
        "wps_ui_observations": {
            "bytes": 4752,
            "sha256": "3c36e2a97253761dd0b48b9a64af3a014a731fad52c2bb1ac24375fdfb4c9ccc",
        },
    }
    process = addendum["process_boundary"]
    assert process["pre_observation_wps_pids"] == [4872, 11388]
    assert process["post_observation_wps_pids"] == [11388]
    assert process["process_termination_command_used"] is False
    assert process["test_owned_window_closed"] is True
    assert "no process was killed" in process["note"]
