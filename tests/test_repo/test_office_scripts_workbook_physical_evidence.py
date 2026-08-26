"""Fail-closed contracts for VIS-2026-07-17-128 workbook physical evidence."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]
GOLDEN = ROOT / "tests" / "fixtures" / "golden"
FIXTURE = GOLDEN / "old_system_official_office_scripts_workbook_batch_semantics.json"
REPORT_NAME = "office-scripts-workbook-xls-ods-physical-matrix-2026-07-17.md"


def _fixture() -> dict[str, object]:
    return json.loads(FIXTURE.read_text(encoding="utf-8"))


def _addendum() -> dict[str, object]:
    return _fixture()["xls_ods_physical_addendum"]


def test_official_sources_and_execution_scope_are_pinned() -> None:
    fixture = _fixture()
    addendum = _addendum()
    assert addendum["evidence_id"] == "VIS-2026-07-17-128"
    assert fixture["source_repository"]["commit"] == ("5d360fa60fb2297a69a039c83e3deb5a97af3939")
    sources = fixture["sources"]
    assert {name: (facts["bytes"], facts["sha256"]) for name, facts in sources.items()} == {
        "email-chart-table.xlsx": (
            11_526,
            "1da5a2011f4f1e6f2e4407f148e814cae47e93d97335c6819d582b6cf433e1be",
        ),
        "hr-schedule.xlsx": (
            13_698,
            "370b8b9b76d8a8c9173de9920ff7fdc92a179874ee0f097cc588e4874293447a",
        ),
        "conditional-formatting-samples.xlsx": (
            16_445,
            "fa17f45f47e0766f13b9cbbf9e63b83962ea7852606bfe0b33807fbfbae5ec64",
        ),
        "task-reminders.xlsx": (
            194_945,
            "b9ad09dfa594fb7397d5a772ed2ed96b865f466680877cf482cc0daaeb516abb",
        ),
    }
    assert addendum["project_heads"] == {
        "docwen-ref-tk": "ec929828",
        "docwen-ref-pyside6": "63db927c",
        "docwen-current": "f98fb9d",
    }


def test_all_conversions_and_same_target_physical_projections_are_exact() -> None:
    execution = _addendum()["execution"]
    assert execution["target_formats"] == ["xls", "ods"]
    assert execution["successful_conversion_count"] == execution["expected_conversion_count"] == 24
    assert execution["successful_excel_readonly_projection_count"] == execution["excel_readonly_projection_count"] == 28
    assert execution["output_pdf_page_count"] == 96
    assert execution["rendered_page_count_including_sources"] == 112
    assert execution["contact_sheet_count"] == 4
    assert execution["same_target_three_project_groups"] == 8
    assert execution["same_target_workbook_semantic_exact_groups"] == 8
    assert execution["same_target_pdf_projection_exact_groups"] == 8
    assert execution["same_target_pdf_text_exact_groups"] == 8

    projections = _addendum()["cross_project_projections"]
    assert set(projections) == {
        "conditional-formatting-samples.xlsx",
        "email-chart-table.xlsx",
        "hr-schedule.xlsx",
        "task-reminders.xlsx",
    }
    assert all(set(targets) == {"xls", "ods"} for targets in projections.values())
    assert sum(target["page_count_each"] for targets in projections.values() for target in targets.values()) == 32
    assert (
        projections["conditional-formatting-samples.xlsx"]["ods"]["pdf_projection_sha256"]
        == "e601467f1e9922e39debe194b74a7e54e528161d2f4b77f0c9c55f6b13a02683"
    )
    assert (
        projections["task-reminders.xlsx"]["ods"]["workbook_semantic_sha256"]
        == "b1b1df6ca29c707d6bb5ef93f3fad2d67e64bfa3a7fb49d27e60afa90223660c"
    )


def test_source_features_are_retained_without_overclaiming_static_sample_scope() -> None:
    retention = _addendum()["source_relative_retention"]
    assert retention["all_eight_target_outputs_retain_source_visible_pdf_text"] is True
    assert retention["all_eight_target_outputs_retain_source_pdf_page_count"] is True

    conditional = retention["conditional-formatting-samples"]
    assert conditional["ods_source_pdf_pixels_exact"] is True
    assert conditional["source_format_condition_count"] == 0
    assert "does not prove conditional-formatting retention" in conditional["boundary"]

    email = retention["email-chart-table"]
    assert (email["source_formula_count"], email["source_table_count"]) == (9, 1)
    assert email["xls_retains_form_control"] is True
    assert email["ods_retains_form_control"] is False
    assert email["formula_result_values_retained"] is True

    hr = retention["hr-schedule"]
    assert (
        hr["source_formula_count"],
        hr["source_table_count"],
        hr["source_hyperlink_count"],
        hr["source_defined_name"],
    ) == (3, 1, 6, "MeetingDuration")
    assert hr["all_targets_retain_formula_cells_table_hyperlink_set_and_defined_name"]

    task = retention["task-reminders"]
    assert (
        task["source_table_count"],
        task["source_hyperlink_count"],
        task["source_merge_count"],
        task["source_picture_count"],
    ) == (1, 11, 2, 1)
    assert task["all_targets_retain_table_hyperlink_set_merges_and_picture"]
