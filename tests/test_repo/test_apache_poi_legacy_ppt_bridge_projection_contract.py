"""Fail-closed contracts for VIS-2026-07-17-130 legacy PPT evidence."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]
GOLDEN = ROOT / "tests" / "fixtures" / "golden"
FIXTURE = GOLDEN / "old_system_pptx_to_markdown_semantics.json"
REPORT_NAME = "apache-poi-legacy-ppt-bridge-matrix-2026-07-17.md"


def _addendum() -> dict[str, object]:
    fixture = json.loads(FIXTURE.read_text(encoding="utf-8"))
    return fixture["legacy_ppt_addendum"]


def test_official_legacy_ppt_sources_and_powerpoint_oracle_are_pinned() -> None:
    addendum = _addendum()
    assert addendum["evidence_id"] == "VIS-2026-07-17-130"
    assert addendum["source_repository"] == "apache/poi"
    assert addendum["source_commit"] == "86e967d9b28d6a322a87ae8fcbf2a7eeb56cef96"
    assert addendum["screened_source_count"] == 5

    sources = addendum["selected_sources"]
    assert set(sources) == {
        "54880_chinese.ppt",
        "basic_test_ppt_file.ppt",
        "pictures.ppt",
        "table_test.ppt",
    }
    assert (sources["54880_chinese.ppt"]["git_blob"], sources["54880_chinese.ppt"]["bytes"]) == (
        "b9830f523ebc14e0b22540989b2c1609455d2ad9",
        103_936,
    )
    assert sources["pictures.ppt"]["sha256"] == ("5318e8486f47c7d9e838eaa4edfe385ba27d7d0ce145364792df53488430e98c")
    assert sources["table_test.ppt"]["git_blob"] == ("d1c688a3ac6f7c221e724a832518ab6c836a5f59")

    physical = addendum["powerpoint_source_oracle"]
    assert (physical["consumer"], physical["version"], physical["build"]) == (
        "Microsoft PowerPoint",
        "16.0",
        "20131",
    )
    assert physical["screened_open_pptx_pdf_success_count"] == physical["screened_expected_count"] == 5
    assert physical["selected_pdf_page_count"] == 9
    assert (physical["contact_sheet_bytes"], physical["contact_sheet_sha256"]) == (
        128_909,
        "712a9d1cf632c3bb368eb10c3988d8019c80d53ac992d8e3f258fc19c4bef20b",
    )


def test_reference_public_route_failure_is_explicit_and_not_promoted_to_pass() -> None:
    addendum = _addendum()
    execution = addendum["public_ppt_route"]
    assert execution["docwen_ref_tk_success_count"] == 0
    assert execution["docwen_ref_pyside6_success_count"] == 0
    assert execution["docwen_current_pre_fix_success_count"] == 4
    assert execution["docwen_current_post_fix_success_count"] == 4
    assert execution["expected_case_count_each"] == 4
    assert execution["current_diagnostic_codes_each"] == ["PPTX2MD-OK", "FINALIZER_DONE"]

    boundary = addendum["reference_environment_boundary"]
    assert boundary["classification"] == "legacy_reference_defect_not_current_regression"
    assert boundary["accepted_as_public_route_success"] is False
    assert "Hiding the application window is not allowed" in boundary["tk_debug_message"]
    assert "Hiding the application window is not allowed" in boundary["old_pyside6_debug_message"]
    assert "fail all four public PPT routes" in boundary["detail"]


def test_current_hub_artifact_defect_is_fixed_at_the_shared_abstraction() -> None:
    addendum = _addendum()
    fix = addendum["current_fix"]
    assert fix["classification"] == "real_current_only_final_artifact_defect_fixed"
    assert fix["pre_fix_picture_final_file_count"] == 1
    assert fix["post_fix_picture_final_file_count"] == 6
    assert "five image links" in fix["before"]
    assert "five exact image artifacts" in fix["after"]
    assert fix["production"] == [
        "packages/core/src/docwen_core/protocols/execution_context.py",
        "packages/core/src/docwen_core/protocols/hub_context.py",
    ]
    assert fix["regression"].endswith("TestPptToMd::test_ppt_to_md_finalizes_images_registered_through_hub_workspace")

    protocol = (ROOT / fix["production"][0]).read_text(encoding="utf-8")
    hub = (ROOT / fix["production"][1]).read_text(encoding="utf-8")
    assert "def registered_artifacts(self) -> list[ArtifactManifest]" in protocol
    assert "return self._delegate.registered_artifacts" in hub


def test_same_wps_hub_downstream_and_current_public_outputs_are_exact() -> None:
    addendum = _addendum()
    bridge = addendum["wps_bridge_oracle"]
    assert bridge["backend"] == "WPS Presentation"
    assert bridge["success_count"] == bridge["expected_success_count"] == 4
    assert len(bridge["pptx_sha256"]) == 4

    matrix = addendum["same_wps_pptx_downstream"]
    assert matrix["success_count"] == matrix["expected_success_count"] == 12
    assert matrix["three_project_normalized_markdown_exact_each"] is True
    assert matrix["three_project_resource_bytes_exact_each"] is True
    assert matrix["current_public_ppt_matches_same_wps_pptx_each"] is True
    assert set(matrix["cases"]) == set(addendum["selected_sources"])
    assert matrix["cases"]["pictures.ppt"]["resource_count_each"] == 5
    assert len(matrix["cases"]["pictures.ppt"]["resource_sha256"]) == 5

    process = addendum["process_boundary"]
    assert process["screen_powerpoint_still_present_after_30_seconds"] is True
    assert process["wps_bridge_wpp_present_immediately_after"] is True
    assert process["post_settlement_only_preexisting_wps_pid"] == 11388
    assert process["harness_termination_command_used"] is False

    classification = addendum["classification"]
    assert classification == {
        "current_only_functional_gap_found": True,
        "current_only_gap_fixed": True,
        "reference_public_route_environment_boundary_accepted_as_success": False,
        "focused_legacy_ppt_parity": "pass_after_fix_with_reference_environment_boundary",
        "broad_presentation_or_overall_parity": "not_proven",
    }
