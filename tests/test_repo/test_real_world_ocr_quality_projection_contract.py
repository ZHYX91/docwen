from __future__ import annotations

import json
from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = pytest.mark.contract

PROJECT_ROOT = Path(__file__).resolve().parents[2]
FIXTURE = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_real_world_ocr_quality_semantics.json"


def _fixture() -> dict:
    return json.loads(FIXTURE.read_text(encoding="utf-8"))


def test_real_world_ocr_fixture_records_reproducible_nonredistributed_source() -> None:
    source = _fixture()["source"]
    assert source["dataset"] == "FUNSD: A Dataset for Form Understanding in Noisy Scanned Documents"
    assert source["download_url"] == "https://guillaumejaume.github.io/FUNSD/dataset.zip"
    assert source["zip_bytes"] == 16838830
    assert source["zip_sha256"] == "c31735649e4f441bcbb4fd0f379574f7520b42286e80b01d80b445649d54761f"
    assert source["actual_image_count"] == 199
    assert source["training_image_count"] == 149
    assert source["testing_image_count"] == 50
    assert source["selected_page_count"] == 5
    assert source["license_boundary"] == "non-commercial_research_and_educational_use_only"
    assert source["redistribution"] == "external_dataset_retained_only_in_workspace_bound_temporary_probe"
    fixture_files = PROJECT_ROOT / "tests" / "fixtures" / "files"
    assert not (fixture_files / "dataset.zip").exists()
    assert not any(fixture_files.rglob("82092117.png"))


def test_real_world_ocr_selection_and_thresholds_were_fixed_and_pass() -> None:
    data = _fixture()
    pages = data["selected_pages"]
    assert [page["id"] for page in pages] == [
        "82092117",
        "82200067_0069",
        "82250337_0338",
        "82251504",
        "82252956_2958",
    ]
    assert data["source"]["selection_rule_fixed_before_execution"].startswith("first five")
    thresholds = data["quality_metric_contract"]["thresholds_fixed_before_results"]
    result = data["three_project_result"]
    assert thresholds == {
        "every_page_nonempty": True,
        "minimum_per_page_normalized_character_f1": 0.6,
        "minimum_micro_normalized_character_f1": 0.75,
        "minimum_micro_exact_token_recall": 0.5,
    }
    assert result["all_projects_pass_fixed_thresholds"] is True
    assert all(page["output_line_count"] > 0 for page in pages)
    assert min(page["normalized_character_f1"] for page in pages) == 0.95246
    assert result["minimum_page_normalized_character_f1"] >= thresholds["minimum_per_page_normalized_character_f1"]
    assert result["micro_normalized_character_metrics"]["f1"] >= thresholds["minimum_micro_normalized_character_f1"]
    assert result["micro_exact_token_metrics"]["recall"] >= thresholds["minimum_micro_exact_token_recall"]


def test_real_world_ocr_three_project_outputs_and_models_are_auditable() -> None:
    data = _fixture()
    execution = data["execution_contract"]
    result = data["three_project_result"]
    assert execution["valid_project_executions"] == 15
    assert execution["all_valid_project_executions_successful"] is True
    assert execution["model_files_byte_identical_across_three_projects"] is True
    assert len(execution["model_sha256"]) == 3
    assert result["all_five_outputs_byte_identical_across_three_projects"] is True
    assert result["combined_output_text_sha256"] == ("f19d1e73f1bd24be11455b6ba4844d0596250f90464cd5538279c6e62506a975")
    assert result["micro_exact_token_metrics"] == {
        "reference": 920,
        "predicted": 862,
        "matched": 748,
        "recall": 0.813043,
        "precision": 0.867749,
        "f1": 0.839506,
    }


def test_real_world_ocr_confidence_policy_regression_has_direct_guard() -> None:
    regression = _fixture()["current_only_regression_fixed"]
    core_source = (PROJECT_ROOT / regression["owner"]).read_text(encoding="utf-8")
    tests = read_source_text(PROJECT_ROOT / "packages" / "core" / "tests" / "test_ocr_*.py")
    assert "confidence_value > 0.5" in core_source
    assert "test_run_ocr_outcome_filters_low_confidence_results" in tests
    assert '("compact nan", "nan")' in tests
    assert "outcome.status is ocr.OcrStatus.SUCCESS" in tests
    assert 'outcome.text == "box trusted\\ncompact trusted"' in tests
    assert regression["real_corpus_output_unchanged_after_fix"] is True
    assert (
        regression["post_fix_combined_output_text_sha256"]
        == _fixture()["three_project_result"]["combined_output_text_sha256"]
    )


def test_real_world_ocr_current_runtime_projection_is_complete() -> None:
    runtime = _fixture()["current_runtime_projection"]
    assert runtime["success"] is True
    assert [(item["kind"], item["name"]) for item in runtime["artifacts"]] == [
        ("primary", "82092117.md"),
        ("image", "82092117.png"),
        ("auxiliary", "82092117_ocr.md"),
    ]
    assert runtime["input_bytes"] == 111080
    assert runtime["output_bytes"] == 112665
    assert runtime["ocr_chars"] == 1256
    assert runtime["diagnostic_codes"] == ["IMG2MD-OK", "IMG2MD-OCR-OK", "FINALIZER_DONE"]
    assert runtime["all_direct_ocr_lines_reachable_in_sidecar"] is True
    assert runtime["workspace_paths_leaked"] is False
    assert len(runtime["invalid_harness_attempts_excluded"]) == 2
