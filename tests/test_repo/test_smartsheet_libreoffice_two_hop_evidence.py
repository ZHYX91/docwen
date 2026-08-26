"""Contracts for the VIS-105 LibreOffice-only SmartSheet two-hop evidence."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
FIXTURE_NAME = "old_system_smartsheet_rich_two_hop_matrix_semantics.json"
REPORT_NAME = "smart-sheet-libreoffice-two-hop-matrix-2026-07-17.md"
FIXTURE_PATH = PROJECT_ROOT / "tests" / "fixtures" / "golden" / FIXTURE_NAME
PROJECTS = {"docwen-ref-tk", "docwen-ref-pyside6", "docwen-current"}


def _addendum() -> dict[str, Any]:
    fixture = json.loads(FIXTURE_PATH.read_text(encoding="utf-8"))
    return fixture["libreoffice_only_two_hop_addendum"]


def test_vis105_pins_official_toolchain_source_and_heads() -> None:
    data = _addendum()

    assert data["vis_id"] == "VIS-2026-07-17-105"
    assert set(data["project_heads"]) == PROJECTS
    assert data["toolchain"] == {
        "version": "LibreOffice 26.2.4.2 0229ac93fcf0d7cbc6376066c6f35021cef002dc",
        "official_msi_sha256": "202F26CDA071C5AA4996A5A28412FDDCEB3891DCEB0366982C62650456C0730F",
        "process_local_administrative_image": "D:/docwen-parity/toolchains/libreoffice-26.2.4/image",
        "system_wide_install_claimed": False,
        "explicit_empty_com_candidate_lists": True,
    }
    assert data["source"]["xlsx_sha256"] == ("8E8B3996091C3120A0252CAB955244025172E8FE9862549F4E1D132F05B9E10E")
    assert data["source"]["generated_xls_sha256"] == (
        "7E0DD558EB467AD875AC6FCB979DCB2A1CE2881303AFA2C460A39F7A22548685"
    )
    assert data["source"]["generated_ods_sha256"] == (
        "9C45738ADBBA71E81660E5E23E131827EB8240DF138919374F1FF48BE87C5E26"
    )


def test_vis105_real_two_hop_matrix_is_libreoffice_only_and_equal() -> None:
    matrix = _addendum()["matrix"]

    assert set(matrix) == {"xls->ods", "ods->xls"}
    for route in matrix.values():
        assert route["attempts"] == route["successes"] == 3
        assert route["backend_per_hop"] == ["LibreOffice", "LibreOffice"]
        assert route["all_project_projections_equal"] is True
        assert route["normalized_projection_sha256"] == (
            "B270046F91B34F284B975DE72E3A8132FA642DEF34823F78989565972509C011"
        )
        assert set(route["final_artifacts"]) == PROJECTS
        assert route["physical_projection"]["page_count_per_project"] == 3
        assert route["physical_projection"]["all_project_pixels_equal"] is True

    xls_artifacts = matrix["ods->xls"]["final_artifacts"].values()
    assert {item["bytes"] for item in xls_artifacts} == {9728}
    assert {item["sha256"] for item in xls_artifacts} == {
        "12A5E2E6024792D6FDFAAA4DBC637DC5915F6B7BD0B3C8842843CAE07A728822"
    }
    ods = matrix["xls->ods"]["ods_package_projection"]
    assert ods == {
        "entry_count": 13,
        "formula_count": 6,
        "validation_count": 1,
        "chart_object_entries": 3,
        "required_text_present": True,
    }


def test_vis105_shared_projection_and_fidelity_boundary_are_explicit() -> None:
    data = _addendum()
    projection = data["shared_projection"]

    assert projection["sheet_names"] == ["Data", "Lookup"]
    assert projection["merged_ranges"] == ["A1:E1"]
    assert projection["formula_values"] == {"E3": 25, "C7": 14, "E7": 95.5}
    assert projection["formula_text"] == {
        "E3": "=C3*D3",
        "C7": "=SUM(C3:C6)",
        "E7": "=SUM(E3:E6)",
    }
    assert projection["list_validation"] == {
        "range": "B3:B6",
        "formula": "Lookup!$A$2:$A$5",
    }
    assert projection["chart_count"] == 1
    assert projection["conditional_format_count"] == 0
    boundary = data["physical_baseline"]["shared_boundary"]
    assert "shifts the chart split" in boundary
    assert "not broad source-fidelity acceptance" in boundary
    assert data["process_evidence"]["no_new_lingering_processes"] is True
