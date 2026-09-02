"""Contracts for VIS-025 SmartSheet CSV/binary artifact evidence."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
FIXTURE_NAME = "old_system_smartsheet_csv_binary_matrix_semantics.json"
REPORT_NAME = "smart-sheet-csv-binary-matrix-2026-07-14.md"
FIXTURE_PATH = PROJECT_ROOT / "tests" / "fixtures" / "golden" / FIXTURE_NAME

EXPECTED_ROUTES = {"csv->xls", "csv->ods", "xls->csv", "ods->csv", "et->csv"}
PROJECTS = {"docwen-ref-tk", "docwen-ref-pyside6", "docwen-current"}


def _fixture() -> dict[str, Any]:
    return json.loads(FIXTURE_PATH.read_text(encoding="utf-8"))


def test_smartsheet_csv_binary_fixture_covers_all_five_routes() -> None:
    fixture = _fixture()

    assert fixture["golden_id"] == "GOLDEN-010"
    assert set(fixture["routes"]) == EXPECTED_ROUTES
    assert set(fixture["matrix"]) == EXPECTED_ROUTES
    for route in fixture["matrix"].values():
        assert set(route) >= PROJECTS
        assert route["docwen-current"]["success"] is True


def test_smartsheet_csv_binary_projection_guards_both_fixed_regressions() -> None:
    fixture = _fixture()
    regressions = fixture["pre_fix_regressions"]
    csv_projection = fixture["binary_to_csv_projection"]

    assert regressions["binary_to_csv_sheet_loss"] == {
        "route": "xls->csv",
        "docwen_ref_tk_artifact_count": 3,
        "docwen_ref_pyside6_artifact_count": 3,
        "docwen_current_artifact_count": 1,
        "current_retained_markers": ["ALPHA_MARKER"],
    }
    assert regressions["csv_to_binary_numeric_cells_became_text"]["current_post_fix_values"] == [
        11,
        22,
        33,
    ]
    assert csv_projection["artifact_count"] == 3
    assert csv_projection["markers"] == ["ALPHA_MARKER", "BETA_MARKER", "HIDDEN_MARKER"]
    assert len(set(csv_projection["sha256"])) == 3


def test_smartsheet_csv_binary_current_runtime_observability_is_complete() -> None:
    current = _fixture()["current_runtime_observability"]

    assert current["representative_route"] == "xls->csv"
    assert current["success"] is True
    assert current["artifact_kinds"] == ["primary", "auxiliary", "auxiliary"]
    assert current["sheet_count_metadata"] == [3, 3, 3]
    assert current["metrics"] == {
        "input_bytes": 19456,
        "output_bytes": 128,
        "sheet_count": 3,
        "total_rows": 6,
    }
    assert current["diagnostics"] == ["SHEETFMT-OK", "FINALIZER_DONE"]


def test_smartsheet_csv_binary_reference_weakness_and_limits_are_explicit() -> None:
    fixture = _fixture()
    old_pyside = fixture["matrix"]["csv->ods"]["docwen-ref-pyside6"]

    assert old_pyside["success"] is False
    assert old_pyside["boundary"] == [
        "Excel COM Open.SaveAs failed",
        "LibreOffice is not installed",
    ]
    assert fixture["matrix"]["csv->ods"]["current_matches_tk_and_exceeds_old_pyside6"] is True
    limits = "\n".join(fixture["accepted_limits"])
    assert "not a final GOLDEN-010 pass" in limits
    assert "LibreOffice is not installed" in limits
