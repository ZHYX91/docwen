"""Contracts for the rich SmartSheet two-hop evidence matrix."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
FIXTURE_NAME = "old_system_smartsheet_rich_two_hop_matrix_semantics.json"
REPORT_NAME = "smart-sheet-rich-two-hop-matrix-2026-07-14.md"
FIXTURE_PATH = PROJECT_ROOT / "tests" / "fixtures" / "golden" / FIXTURE_NAME

EXPECTED_ROUTES = {
    "xls->ods",
    "xls->et",
    "ods->xls",
    "ods->et",
    "et->xls",
    "et->ods",
}
PROJECTS = {"docwen-ref-tk", "docwen-ref-pyside6", "docwen-current"}


def _fixture() -> dict[str, Any]:
    return json.loads(FIXTURE_PATH.read_text(encoding="utf-8"))


def test_smartsheet_rich_two_hop_matrix_covers_all_six_non_xlsx_routes() -> None:
    fixture = _fixture()
    matrix = fixture["matrix"]

    assert fixture["golden_id"] == "GOLDEN-010"
    assert set(fixture["routes"]) == EXPECTED_ROUTES
    assert set(matrix) == EXPECTED_ROUTES
    for route in matrix.values():
        assert route["all_project_projections_equal"] is True
        assert route["retained_rich_projection"] is True
        assert set(route["projects"]) == PROJECTS
        assert route["projects"]["docwen-current"]["cli_success"] is True
        assert route["projects"]["docwen-current"]["output_name"].startswith("vis024-rich.")


def test_smartsheet_rich_projection_and_ods_package_surface_are_guarded() -> None:
    fixture = _fixture()
    projection = fixture["normalized_rich_projection"]
    ods = fixture["ods_output_package_projection"]

    assert projection["sheet_names"] == ["Data", "Lookup", "_Meta"]
    assert projection["sheet_visibility"]["_Meta"] == 0
    assert projection["merge_area"] == "$A$1:$E$1"
    assert projection["representative_cells"]["E7"] == {
        "value": 95.75,
        "formula": "=SUM(E3:E6)",
        "bold": True,
        "italic": True,
        "fill": 13431551,
        "number_format_semantic": "$0.00",
    }
    assert projection["chart_count"] == 1
    assert projection["shape_count"] == 3
    assert projection["hyperlink_count"] == 1
    assert projection["comment_count"] == 1
    assert projection["validation_type_b3"] == 3
    assert projection["auto_filter_mode"] is True
    assert ods["applies_to"] == ["xls->ods", "et->ods"]
    assert ods["all_projects"]["entry_count"] == 8
    assert ods["all_projects"]["formula_count"] == 6
    assert ods["all_projects"]["has_content_validations"] is True
    assert ods["all_projects"]["has_embedded_chart_object"] is True
    assert ods["all_projects"]["media"]["media/image1.png"] == (
        "8d45b06f6c8ea81217ef1d660565828afa4a74b42ddf749c3c1e63caa21a09ff"
    )


def test_smartsheet_current_runtime_observability_survives_finalization() -> None:
    current = _fixture()["current_runtime_observability"]

    assert current["representative_route"] == "xls->ods"
    assert current["success"] is True
    assert current["final_name"] == current["suggested_name"] == "vis024-rich.ods"
    assert current["metadata"] == {
        "source_format": "xls",
        "target_format": "ods",
        "backend": "WPS Spreadsheets -> Microsoft Excel",
    }
    assert current["metrics"] == {
        "input_bytes": 31744,
        "output_bytes": 9074,
        "backend": "WPS Spreadsheets -> Microsoft Excel",
    }
    assert current["diagnostics"] == ["SHEETFMT-OK", "FINALIZER_DONE"]
    assert current["process_remaining"] == []
