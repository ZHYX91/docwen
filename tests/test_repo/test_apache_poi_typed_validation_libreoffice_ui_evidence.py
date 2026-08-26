from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract
ROOT = Path(__file__).resolve().parents[2]
FIXTURE = ROOT / "tests/fixtures/golden/old_system_apache_poi_typed_validation_semantics.json"


def _addendum() -> dict:
    data = json.loads(FIXTURE.read_text(encoding="utf-8"))
    return data["libreoffice_ods_physical_ui_addendum"]


def test_libreoffice_ods_list_rule_is_exact_across_projects() -> None:
    addendum = _addendum()
    assert addendum["stage"] == "VIS-2026-07-17-119"
    assert addendum["source_final_artifact_stage"] == "VIS-2026-07-17-117"
    assert addendum["source_wps_boundary_stage"] == "VIS-2026-07-17-118"
    ui = addendum["physical_libreoffice_ods"]
    assert ui["projects_observed"] == [
        "docwen-ref-tk",
        "docwen-ref-pyside6",
        "docwen-current",
    ]
    b2 = ui["b2_inline_list"]
    assert b2["dropdown_values"] == ["IN", "US", "UK"]
    assert (b2["input_prompt_title"], b2["input_prompt_message"]) == (
        "Country Code Selection",
        "Choose a Country Code.",
    )
    assert (b2["invalid_error_title"], b2["invalid_error_message"]) == (
        "Invalid Country Code",
        "The specified country code is not a valid option.",
    )
    assert b2["invalid_value_rejected"] is True
    assert b2["post_rejection_value"] == "IN"
    assert b2["project_passes"] == 3
    assert b2["cross_project_behavior_exact"] is True


def test_libreoffice_ods_integer_rule_preserves_parity_without_overclaim() -> None:
    b6 = _addendum()["physical_libreoffice_ods"]["b6_whole_number_greater_than_zero"]
    assert (b6["input_prompt_title"], b6["input_prompt_message"]) == (
        "Integer Input",
        "Enter an integer value.",
    )
    assert b6["invalid_keyboard_value"] == 0
    assert (b6["invalid_error_title"], b6["invalid_error_message"]) == (
        "LibreOffice Calc",
        "无效的值。",
    )
    assert b6["invalid_value_rejected"] is True
    assert b6["post_rejection_value"] == 1
    assert b6["project_passes"] == 3
    assert b6["cross_project_behavior_exact"] is True
    assert b6["source_custom_error_preserved"] is False
    assert "not a current-only regression" in b6["classification"]
    assert "not" in b6["classification"] and "accepted" in b6["classification"]


def test_literal_text_diagnostic_is_excluded_fail_closed() -> None:
    boundary = _addendum()["interaction_evidence_boundary"]
    assert boundary["authoritative_invalid_entry"] == ("individual keyboard events followed by Return")
    assert boundary["literal_text_injection_used_for_verdict"] is False
    assert "bypass validation" in boundary["excluded_diagnostic"]
    assert boundary["classification"].startswith("fail-closed")


def test_libreoffice_environment_cleanup_and_external_identity_are_frozen() -> None:
    addendum = _addendum()
    environment = addendum["environment"]
    assert environment["application"] == "LibreOffice Calc"
    assert environment["version"].startswith("26.2.4.2 ")
    assert environment["bootstrap_user_installation"] == "$ORIGIN/../Data/settings"
    assert environment["system_wide_installation_claimed"] is False
    assert environment["online_service_used"] is False

    process = addendum["process_boundary"]
    assert process["pre_observation_soffice_processes"] == []
    assert process["post_observation_soffice_processes"] == []
    assert process["process_termination_command_used"] is False
    assert process["private_profile_used"] is True

    ui = addendum["physical_libreoffice_ods"]
    assert ui["all_test_copies_closed_without_save"] is True
    assert ui["all_post_observation_sizes_and_hashes_unchanged"] is True

    evidence = addendum["external_evidence"]
    assert (evidence["files"], evidence["bytes"]) == (4, 23531)
    assert evidence["libreoffice_ui_observations"] == {
        "bytes": 5494,
        "sha256": "f15ffa99ac6afa4d8f5eada2c336c10d56053a89d5a6c9ec5e4a8f1ce80dded0",
    }
