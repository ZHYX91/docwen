from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

PROJECT_ROOT = Path(__file__).resolve().parents[2]
FIXTURE = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_real_world_chinese_photo_ocr_semantics.json"


def _fixture() -> dict:
    return json.loads(FIXTURE.read_text(encoding="utf-8"))


def test_chinese_photo_ocr_source_is_reproducible_and_not_redistributed() -> None:
    source = _fixture()["source"]
    assert source["dataset"] == "CCPD2020 / CCPD-Green"
    assert source["archive_bytes"] == 907711344
    assert source["archive_md5"] == "eb93e88c5988879f8da6f92bbf083324"
    assert source["archive_jpeg_count"] == 11776
    assert source["selected_photo_count"] == 5
    assert source["selection_rule_fixed_before_execution"].startswith("first five JPEG entries")
    assert source["license_boundary"].startswith("CC-BY-4.0")
    assert source["redistribution"] == "external_dataset_retained_only_in_workspace_bound_temporary_probe"
    fixture_files = PROJECT_ROOT / "tests" / "fixtures" / "files"
    assert not any(fixture_files.rglob("CCPD2020.zip"))
    assert not any(fixture_files.rglob("ccpd2020_test_*.jpg"))
    assert not any(fixture_files.rglob("ccpd2020_test_*.png"))


def test_chinese_photo_ocr_selection_and_thresholds_are_frozen() -> None:
    data = _fixture()
    assert data["execution_contract"]["contract_created_before_any_docwen_ocr_execution"] is True
    selected = data["selected_inputs"]
    assert [item["label"] for item in selected] == [
        "皖AD66178",
        "皖AD32818",
        "皖AF01060",
        "皖AD67887",
        "皖ADB0129",
    ]
    assert [item["crop_size"] for item in selected] == [
        [60, 23],
        [67, 24],
        [76, 23],
        [75, 28],
        [75, 27],
    ]
    assert data["quality_metric_contract"]["thresholds_fixed_before_results"] == {
        "all_three_projects_outputs_byte_identical_per_input": True,
        "full_frame": {
            "every_page_nonempty": True,
            "minimum_pages_with_best_line_similarity_at_least_0_50": 3,
            "minimum_mean_best_line_similarity": 0.55,
            "minimum_pages_with_province_character_present": 3,
        },
        "bbox_crop": {
            "every_page_nonempty": True,
            "minimum_per_page_best_line_similarity": 0.5,
            "minimum_mean_best_line_similarity": 0.75,
            "minimum_exact_label_matches": 2,
        },
    }
    assert data["three_project_result"]["thresholds_changed_after_results"] is False


def test_chinese_photo_ocr_records_identity_and_failed_quality_without_overclaim() -> None:
    data = _fixture()
    execution = data["execution_contract"]
    result = data["three_project_result"]
    assert execution["valid_project_executions"] == 30
    assert execution["all_valid_project_executions_successful"] is True
    assert execution["model_files_byte_identical_across_three_projects"] is True
    assert len(execution["model_sha256"]) == 3
    assert result["all_ten_outputs_byte_identical_across_three_projects"] is True
    assert result["canonical_outputs_sha256"] == ("14403e20d43940c362b6a2caa98519510acb76d31c6330babf6f6c49fd786db5")
    assert result["identity_threshold_passed"] is True
    assert result["all_quality_thresholds_passed"] is False
    assert result["full_frame"] == {
        "every_page_nonempty": False,
        "nonempty_pages": 4,
        "pages_with_best_line_similarity_at_least_0_50": 2,
        "mean_best_line_similarity": 0.325,
        "pages_with_province_character_present": 2,
    }
    assert result["bbox_crop"]["nonempty_pages"] == 0
    assert result["bbox_crop"]["exact_label_matches"] == 0


def test_chinese_photo_ocr_current_runtime_projection_is_complete() -> None:
    runtime = _fixture()["current_runtime_projection"]
    assert runtime["success"] is True
    assert [(item["kind"], item["name"]) for item in runtime["artifacts"]] == [
        ("primary", "ccpd2020_test_00.md"),
        ("image", "ccpd2020_test_00.jpg"),
        ("auxiliary", "ccpd2020_test_00_ocr.md"),
    ]
    assert [item["sha256"] for item in runtime["artifacts"]] == [
        "ef59ba6fdbe72a2ab128b672542842543209175f3b8994a3500ac890f84bd863",
        "bb0df650fc1a7e5cf30cac336c02e503658648a0e5713877b7b57ed2a6f23634",
        "e8d767eb9eeff5069d38604efc45f9a593bef0ae10736d3d63add18ac6cbe4e0",
    ]
    assert runtime["input_bytes"] == 98705
    assert runtime["output_bytes"] == 99017
    assert runtime["ocr_chars"] == 16
    assert runtime["diagnostic_codes"] == ["IMG2MD-OK", "IMG2MD-OCR-OK", "FINALIZER_DONE"]
    assert runtime["all_direct_ocr_lines_reachable_in_sidecar"] is True
    assert runtime["primary_references_sidecar"] is True
    assert runtime["sidecar_references_retained_image"] is True
    assert runtime["workspace_paths_leaked"] is False


def test_chinese_photo_ocr_classifies_shared_boundary_not_current_regression() -> None:
    data = _fixture()
    classification = data["classification"]
    assert classification["current_only_functional_gap_found"] is False
    assert classification["route_or_option_wiring_gap_found"] is False
    assert classification["shared_upstream_model_quality_boundary_found"] is True
    assert classification["production_change_made"] is False
    assert len(data["known_limits"]) == 5
    assert any("not an accepted difference" in item for item in data["known_limits"])
    assert any("NOT PASSED YET" in item for item in data["known_limits"])
