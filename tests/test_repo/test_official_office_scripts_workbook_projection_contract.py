from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

PROJECT_ROOT = Path(__file__).resolve().parents[2]
FIXTURE = PROJECT_ROOT / "tests/fixtures/golden/old_system_official_office_scripts_workbook_batch_semantics.json"


def _fixture() -> dict:
    return json.loads(FIXTURE.read_text(encoding="utf-8"))


def test_office_scripts_sources_are_official_frozen_and_not_distributed() -> None:
    data = _fixture()
    repository = data["source_repository"]
    assert repository["owner"] == "OfficeDev"
    assert repository["repository"] == "office-scripts-docs"
    assert repository["commit"] == "5d360fa60fb2297a69a039c83e3deb5a97af3939"
    assert repository["features_screened_before_docwen_execution"] is True
    expected = {
        "email-chart-table.xlsx": (11526, "1da5a2011f4f1e6f2e4407f148e814cae47e93d97335c6819d582b6cf433e1be"),
        "hr-schedule.xlsx": (13698, "370b8b9b76d8a8c9173de9920ff7fdc92a179874ee0f097cc588e4874293447a"),
        "conditional-formatting-samples.xlsx": (
            16445,
            "fa17f45f47e0766f13b9cbbf9e63b83962ea7852606bfe0b33807fbfbae5ec64",
        ),
        "task-reminders.xlsx": (194945, "b9ad09dfa594fb7397d5a772ed2ed96b865f466680877cf482cc0daaeb516abb"),
    }
    assert {name: (source["bytes"], source["sha256"]) for name, source in data["sources"].items()} == expected
    fixture_files = PROJECT_ROOT / "tests/fixtures/files"
    for name in expected:
        assert not any(fixture_files.rglob(name))


def test_office_scripts_feature_screen_prevents_filename_overclaim() -> None:
    sources = _fixture()["sources"]
    email = sources["email-chart-table.xlsx"]
    assert email["formula_count"] == email["formula_with_cached_value_count"] == 9
    assert email["formula_samples"]["D9"] == {
        "formula": "=SUBTOTAL(109,Table1[Amount])",
        "cached": 1291,
    }
    assert email["chart_count"] == 0

    hr = sources["hr-schedule.xlsx"]
    assert hr["formula_count"] == hr["formula_with_cached_value_count"] == 3
    assert hr["hyperlink_count"] == 6

    conditional = sources["conditional-formatting-samples.xlsx"]
    assert len(conditional["sheet_names"]) == 8
    assert conditional["conditional_formatting_range_count"] == 0
    assert "accompanying Office Script adds" in conditional["boundary"]

    task = sources["task-reminders.xlsx"]
    assert (task["image_count"], task["merge_count"], task["hyperlink_count"]) == (1, 2, 11)
    assert all(source["data_validation_count"] == 0 for source in sources.values())
    assert all(source["chart_count"] == 0 for source in sources.values())


def test_office_scripts_three_project_projections_are_equal() -> None:
    data = _fixture()
    execution = data["execution_contract"]
    assert execution["valid_three_project_executions"] == 12
    assert execution["all_executions_successful"] is True
    assert execution["office_processes_started"] is False
    expected = {
        "email-chart-table.xlsx": (10, 1, "89c279818d262cb572d1f9592edc983a470eb3da0aa2b30bdd15dcd6f3e9bcd1"),
        "hr-schedule.xlsx": (8, 2, "3cc010ad6c4955eec82eb6cd78aed3cb2aec1e28531da37803f0a006171a2c61"),
        "conditional-formatting-samples.xlsx": (
            58,
            8,
            "0b378b78e75fee04ce6582c281d6a0e9b1f63e59ff7f93bb75141cade2bbf174",
        ),
        "task-reminders.xlsx": (15, 1, "08ca1c345240846fbfe2bb35bf63be22f4b06388b7a76a77cc585e67310aab64"),
    }
    for name, (entries, headings, digest) in expected.items():
        projection = data["normalized_projections"][name]
        assert projection["all_three_projects_equal"] is True
        assert projection["entries"] == entries
        assert projection["headings"] == headings
        assert projection["projection_sha256"] == digest
        assert projection["old_pyside6_current_raw_equal"] is True


def test_official_cached_formula_policy_is_recorded_without_formula_text_claim() -> None:
    contract = _fixture()["cached_formula_contract"]
    assert contract["source_formula_cells"] == 12
    assert contract["source_formula_cells_with_cached_values"] == 12
    assert contract["all_three_projects_emit_cached_values_not_formula_text"] is True
    assert contract["email_cached_tokens"] == ["124.8", "279.5", "1291", "978.1"]
    assert contract["hr_cached_tokens"] == [
        "2021-06-05 09:00:00",
        "2021-06-05 12:00:00",
        "2021-06-05 15:00:00",
    ]
    assert "not formula-text preservation" in contract["classification"]


def test_official_workbook_current_runtime_projection_is_complete() -> None:
    runtime = _fixture()["current_runtime_finalizer"]
    assert runtime["success"] is True
    assert runtime["artifact_count"] == 1
    assert runtime["primary_name"] == "email-chart-table.md"
    assert runtime["primary_bytes"] == runtime["output_bytes"] == 1061
    assert runtime["primary_sha256"] == ("fa5de70240390d161a64ba45f65aae7527fa546052f6ad775d1907c61574611d")
    assert runtime["input_bytes"] == 11526
    assert (runtime["sheet_count"], runtime["row_count"], runtime["column_count"]) == (1, 9, 6)
    assert (runtime["block_count"], runtime["image_count"]) == (1, 0)
    assert runtime["diagnostics"] == ["SHEET2MD-OK", "FINALIZER_DONE"]
    assert runtime["workspace_paths_leaked"] is False
