"""Contracts for VIS-026 remaining SmartDoc route evidence."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
FIXTURE_NAME = "old_system_smartdoc_remaining_route_matrix_semantics.json"
REPORT_NAME = "smartdoc-remaining-route-matrix-2026-07-14.md"
FIXTURE_PATH = PROJECT_ROOT / "tests" / "fixtures" / "golden" / FIXTURE_NAME
SUPPORTED = {"doc->odt", "doc->rtf", "odt->doc", "rtf->doc", "wps->doc", "wps->odt", "wps->rtf"}
ENHANCEMENTS = {"doc->wps", "odt->wps", "rtf->wps"}
PROJECTS = {"docwen-ref-tk", "docwen-ref-pyside6", "docwen-current"}


def _fixture() -> dict[str, Any]:
    return json.loads(FIXTURE_PATH.read_text(encoding="utf-8"))


def test_remaining_smartdoc_fixture_completes_focused_twenty_route_breadth() -> None:
    fixture = _fixture()

    assert fixture["golden_id"] == "GOLDEN-009"
    assert set(fixture["routes"]) == SUPPORTED | ENHANCEMENTS
    assert set(fixture["matrix"]) == SUPPORTED | ENHANCEMENTS
    assert len(fixture["coverage_context"]["previously_evidenced_routes"]) == 10
    assert fixture["coverage_context"]["focused_real_route_breadth_after_this_fixture"] == "20/20"


def test_supported_remaining_smartdoc_routes_match_source_across_projects() -> None:
    matrix = _fixture()["matrix"]

    for route_name in SUPPORTED:
        route = matrix[route_name]
        assert route["classification"] == "three-project-supported"
        assert set(route) >= PROJECTS
        assert all(route[project]["success"] is True for project in PROJECTS)
        assert all(route[project]["artifact_count"] == 1 for project in PROJECTS)
        assert route["all_project_projections_equal"] is True
        assert route["all_project_projections_match_source"] is True


def test_old_wps_targets_are_explicit_current_enhancements() -> None:
    matrix = _fixture()["matrix"]

    for route_name in ENHANCEMENTS:
        route = matrix[route_name]
        assert route["classification"] == "current-enhancement"
        for old_project in ("docwen-ref-tk", "docwen-ref-pyside6"):
            assert route[old_project] == {
                "success": False,
                "artifact_count": 0,
                "returned_none": True,
            }
        assert route["docwen-current"]["success"] is True
        assert route["docwen-current"]["artifact_count"] == 1
        assert route["current_projection_matches_source"] is True


def test_remaining_smartdoc_projection_and_runtime_contract_are_guarded() -> None:
    fixture = _fixture()
    projection = fixture["normalized_source_projection"]
    runtime = fixture["current_runtime_contract"]

    assert projection["all_four_source_formats_equal"] is True
    assert projection["styles"] == {
        "VIS-026 Remaining SmartDoc Matrix": "Title",
        "Semantic Heading": "Heading 1",
        "BODY_MARKER": "Normal",
    }
    assert projection["bold_marker_direct"] is True
    assert projection["italic_marker_direct"] is True
    assert projection["table"] == [["Item", "Qty", "Note"], ["Alpha", "7", "MATRIX_MARKER"]]
    assert projection["sections"] == ["portrait", "landscape"]
    assert projection["media_count"] == 1
    assert fixture["container_projection"]["all_successful_artifacts_match_their_target_container"] is True
    assert runtime == {
        "all_successful": True,
        "all_names_source_owned": True,
        "all_metadata_matches_route_endpoints": True,
        "all_diagnostics": ["DOCX-SMARTDOC-OK", "FINALIZER_DONE"],
        "all_metrics_match_final_input_and_output_bytes": True,
        "private_docx_hubs_exposed_as_final_artifacts": False,
    }


def test_remaining_smartdoc_process_and_limits_do_not_overclaim() -> None:
    fixture = _fixture()

    assert [item["pid"] for item in fixture["process_evidence"]["pre_existing"]] == [1796, 18764, 20176]
    assert fixture["process_evidence"]["new_after_final_settle"] == []
    limits = "\n".join(fixture["accepted_limits"])
    assert "not a final GOLDEN-009 pass" in limits
    assert "LibreOffice is not installed" in limits
    assert "Broader real-world document batches" in limits
