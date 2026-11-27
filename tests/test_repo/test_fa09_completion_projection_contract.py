"""Fail-closed evidence guards for VIS-192 / finite-contract FA-09 closure."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = [pytest.mark.contract, pytest.mark.golden]

PROJECT_ROOT = Path(__file__).resolve().parents[2]
PROJECTION_SHA = "1ADD0084CA59296E764D900A091D880E139664B2226EE068EA3D050431CD2E0C"


def test_fa09_completion_fixture_records_current_green_and_exact_reference_defects() -> None:
    path = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_pdf_operations_semantics.json"
    fixture = json.loads(path.read_text(encoding="utf-8"))
    addendum = fixture["fa09_completion_addendum"]

    assert addendum["stage_id"] == "VIS-2026-07-23-192"
    assert addendum["contract_slots"] == 27
    assert addendum["new_n1_n2_slots"] == 15
    assert addendum["current_new_slots_green"] == 5
    assert addendum["current_custom_split_internal_goto_targets"] == {"S1": "S3", "S2": "S4"}
    assert addendum["reference_custom_split_internal_goto_targets"] == {
        "docwen-ref-tk": {},
        "docwen-ref-pyside6": {},
    }
    assert addendum["reference_n1_defect"]["current_same_routes_green"] is True
    assert addendum["n2_png"] == {
        "pages_per_project": 5,
        "geometry": [1241, 1754],
        "dpi": 150,
        "decoded_pixels_equal_across_projects": True,
    }
    assert addendum["n2_markdown"]["frozen_anchors_green_per_project"] == 10
    assert addendum["n2_markdown"]["current_page_resources"] == 5
    assert addendum["n2_markdown"]["docwen_ref_tk_page_resources"] == 2
    assert addendum["n2_markdown"]["docwen_ref_pyside6_page_resources"] == 2
    assert addendum["current_n1_markdown"]["frozen_anchors_green"] == 9
    assert addendum["external_evidence"]["matrix_projection_sha256"] == PROJECTION_SHA
    assert addendum["user_accepted_boundary"] is False
    assert addendum["fa09_status"] == "FIXED_AND_VERIFIED"
    assert addendum["overall_status"] == "NOT_PASSED_YET"
