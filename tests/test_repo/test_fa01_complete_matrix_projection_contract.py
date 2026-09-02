"""Fail-closed evidence guards for VIS-208 / finite-contract FA-01 closure."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = [pytest.mark.contract, pytest.mark.golden]
ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "fa01-complete-matrix-artifact-oracle-2026-07-24.md"
CARD_NAME = "fa01-complete-matrix-stage-card-2026-07-24.md"
FIXTURE_NAME = "current_fa01_complete_matrix_semantics.json"
FIXTURE = ROOT / "tests/fixtures/golden" / FIXTURE_NAME
STATUS = "PASS_WITH_USER_ACCEPTED_BOUNDARY"
SELECTION = "FA-01=TEMPLATE-FREEZE-B"
FORMULA = "=COUNTA(A5:E7)"


def _data() -> dict:
    return json.loads(FIXTURE.read_text(encoding="utf-8"))


def test_fa01_identity_accounting_and_disposition_are_exact() -> None:
    data = _data()
    assert data["case_id"] == "current_fa01_complete_matrix_semantics"
    assert data["stage_id"] == "VIS-2026-07-24-208"
    assert data["status"] == STATUS
    assert data["overall_parity"] == "NOT_PASSED_YET"
    assert data["user_selection"] == SELECTION
    assert data["stage_start"] == {
        "parent": "b761a1dcb2d6412d84fd5a5985c81174e4b2ef9c",
        "tree": "900bb6957eb4a7b22156fc9864708b9fb71008d9",
    }
    assert data["accounting"] == {
        "family_logical_outputs": 18,
        "reused_b1_outputs": 9,
        "fresh_n1_outputs": 9,
        "fresh_n1_by_project": {"current": 3, "tk": 3, "pyside6": 3},
    }


def test_fa01_corrected_source_template_and_formula_oracle_are_exact() -> None:
    data = _data()
    assert data["frozen_inputs"] == {
        "markdown": {
            "bytes": 21005,
            "sha256": "d78a2bd8483184f8d9b3c3100b9a3b6d8dad6ea6569b63d2d22a916bf7a41d39",
        },
        "image": {
            "bytes": 141190,
            "sha256": "5baa96d93e1f611d5bac1ca2379302049b2673d9d57aa2ff0582b51767028c31",
        },
        "corrected_template": {
            "bytes": 4434,
            "sha256": "282b3cfdbd76235cf5e8ddb7ef3eabcd7c915187f875a1accbeb84f31f3e40bb",
            "formula": FORMULA,
            "source_data_rows": 3,
            "formula_result_in_excel": 15,
        },
    }
    for project in ("current", "tk", "pyside6"):
        workbook = data["projects"][project]["xlsx"]
        assert workbook["formula"] == FORMULA
        assert workbook["image_exact"] is True
        assert workbook["freeze_panes"] is None


def test_fa01_project_artifact_and_office_projection_is_pinned() -> None:
    data = _data()
    assert {
        project: {kind: payload["sha256"] for kind, payload in artifacts.items()}
        for project, artifacts in data["projects"].items()
    } == {
        "current": {
            "docx": "c52e6c297a4215463f15e2a793939ad8c13b4c87fc71c4c03f0c33046b6e7c08",
            "xlsx": "874aeefb7c1b8e43a47ea7355524593b3525ace3f826a1b96d115e045386ac89",
            "csv": "6d6d5b9ffea26e9ce423fbf6ee063185b610390d13465ad5fd45d389def9c26a",
        },
        "tk": {
            "docx": "0a8be72576957be75901f43a26e2adf608ceb3f3849a381dbc048f457584d6e7",
            "xlsx": "731476f0f0ccf11f84fee3603639feee5df5b293a2d42370d5421638f5bfb4a4",
            "csv": "9178ea07541721a438aecdce863e41bbc09ab6f51ade2564ef4c791b57af7cc5",
        },
        "pyside6": {
            "docx": "56682ccdaa529b3ad1265ce87d3b8a11568f296022947ba49e8931ec476c0b21",
            "xlsx": "f9a750008abe18fad129e1874b01b3b2653f6ffd223289d7cc7a4df64e02e800",
            "csv": "9178ea07541721a438aecdce863e41bbc09ab6f51ade2564ef4c791b57af7cc5",
        },
    }
    assert data["office"] == {
        "word_docx_pages": {"current": 14, "tk": 17, "pyside6": 17},
        "wps_docx_pages": {"current": 14, "tk": 17, "pyside6": 17},
        "excel_pdf_pages_per_project": 3,
        "excel_cross_project_render_pixels_equal": True,
        "rendered_pdfs": 9,
        "rendered_pages": 105,
        "contact_sheets_manually_inspected": 9,
        "preexisting_wps_pids": [6496, 13884],
        "process_delta": 0,
    }


def test_fa01_missing_freeze_is_the_only_accepted_defect() -> None:
    data = _data()
    assert data["dispositions"]["accepted_boundary"] == [
        "N1 template and all three derived XLSX artifacts omit the A5 freeze pane"
    ]
    assert set(data["dispositions"]["not_accepted"]) == {
        "missing, invalid or empty artifact",
        "source or template mutation",
        "mandatory content loss",
        "wrong formula or formula result",
        "wrong or missing source image",
        "path leak",
        "Office repair prompt",
        "process residue",
        "undisclosed defect",
    }
    assert data["mandatory_predicates"] == {
        "all_18_outputs_accounted": True,
        "all_9_fresh_routes_succeeded": True,
        "all_packages_valid": True,
        "source_and_templates_immutable": True,
        "docx_mandatory_anchors_complete": True,
        "docx_source_image_exact": True,
        "xlsx_three_rows_formula_merge_image_exact": True,
        "csv_three_rows_exact": True,
        "excel_formula_value_15": True,
        "word_and_wps_open_without_repair": True,
        "current_only_regression_remaining": False,
        "path_leak": False,
        "process_residue": False,
    }
    assert data["external_evidence"] == {
        "root": "D:/docwen-parity/vis208-fa01-complete-b761a1d-v1",
        "carrier_files_excluding_dependency_junction": 158,
        "carrier_bytes_excluding_dependency_junction": 28338339,
        "final_oracle_bytes": 5007,
        "final_oracle_sha256": "9759bb57822ee941f4bd5da66016c711dfb0976e0ca05f93cb791dc256a44f9d",
        "binaries_checked_in": False,
    }
