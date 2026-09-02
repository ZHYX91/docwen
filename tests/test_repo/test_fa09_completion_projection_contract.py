"""Fail-closed evidence guards for VIS-192 / finite-contract FA-09 closure."""

from __future__ import annotations

import json
from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = [pytest.mark.contract, pytest.mark.golden]

PROJECT_ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "fa09-complete-matrix-and-internal-goto-closure-2026-07-23.md"
STAGE_CARD = "fa09-complete-matrix-and-internal-goto-stage-card-2026-07-23.md"
SOURCE_SHA = "55192F12F6AFF1294EB9F40000F9455B5F905D5F661AA308EBFAFE8A31154D02"
SCAN_SHA = "C1C8AAA7267E961404F18455742F0C87A08D3B9930363B05510C08369AA6688F"
PROJECTION_SHA = "1ADD0084CA59296E764D900A091D880E139664B2226EE068EA3D050431CD2E0C"


def _read(path: Path) -> str:
    return read_source_text(path)


def test_fa09_completion_fixture_records_current_green_and_exact_reference_defects() -> None:
    fixture = json.loads(
        _read(PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_pdf_operations_semantics.json")
    )
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


def test_fa09_completion_repairs_have_direct_executable_guards() -> None:
    splitter = _read(
        PROJECT_ROOT
        / "packages"
        / "plugins"
        / "layout"
        / "src"
        / "docwen_plugin_layout"
        / "operations"
        / "converter.py"
    )
    markdown = _read(
        PROJECT_ROOT
        / "packages"
        / "plugins"
        / "layout"
        / "src"
        / "docwen_plugin_layout"
        / "to_markdown"
        / "converter.py"
    )
    tests = _read(PROJECT_ROOT / "packages" / "plugins" / "layout" / "tests" / "test_layout_conversions_*.py")

    assert "def _copy_pages_preserving_internal_gotos" in splitter
    assert "target_source_index not in destination_by_source" in splitter
    assert "def _ocr_page_outcomes" in markdown
    assert "if len(page_outcomes) != physical_page_count:" in markdown
    assert '"fragment_kind": "page"' in markdown
    assert '"page_count": physical_page_count' in markdown
    assert "test_pdf_custom_split_does_not_crosslink_omitted_goto_target" in tests
    assert "test_pdf_to_md_physical_page_matrix_keeps_p_and_k_independent" in tests
    assert "test_pdf_to_md_page_fragment_preserves_ocr_text_exactly" in tests
