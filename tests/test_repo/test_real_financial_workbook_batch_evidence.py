from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

PROJECT_ROOT = Path(__file__).resolve().parents[2]
FIXTURE = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_real_financial_workbook_batch_semantics.json"


def _fixture() -> dict:
    return json.loads(FIXTURE.read_text(encoding="utf-8"))


def test_financial_workbook_sources_are_official_frozen_and_not_distributed() -> None:
    sources = _fixture()["sources"]
    financial = sources["financial_sample"]
    contoso = sources["contoso_pnl"]
    assert financial["publisher"] == "Microsoft"
    assert financial["size_bytes"] == 83418
    assert financial["sha256"] == ("c3f17156ab7c192571ecfc742e88c16a4b7243a4d4b4a420fcb68dec44e10196")
    assert financial["rows"] == 701
    assert financial["columns"] == 16
    assert financial["nonempty_cells"] == 11216
    assert contoso["publisher"] == "Microsoft"
    assert contoso["archive_size_bytes"] == 6592017
    assert contoso["archive_sha256"] == ("a4e8ab694927bd8e394a5c18bc06349093cd96452c5195bffddcc5192d6af65c")
    assert contoso["workbook_size_bytes"] == 7361732
    assert contoso["workbook_sha256"] == ("d9a77819ed43a93cda82ab7ba08dd6406981d9d3eae7cfb7cae4720a40f87ac5")
    assert contoso["ooxml_package_entries"] == 260
    assert len(contoso["sheet_names"]) == 9
    assert sources["redistribution"].startswith("external_official_binaries")
    fixture_files = PROJECT_ROOT / "tests" / "fixtures" / "files"
    assert not any(fixture_files.rglob("Financial Sample.xlsx"))
    assert not any(fixture_files.rglob("ContosoPnL_Excel2013.zip"))
    assert not any(fixture_files.rglob("ContosoPnL_Excel2013.xlsx"))


def test_dense_financial_sample_projection_is_equal_across_production_paths() -> None:
    data = _fixture()
    execution = data["execution_contract"]
    projection = data["financial_sample_projection"]
    assert execution["valid_project_executions_successful"] is True
    assert execution["office_processes_started"] is False
    assert projection["all_four_paths_semantically_equal"] is True
    assert projection["entries"] == 702
    assert projection["headings"] == 1
    assert projection["table_rows_including_header"] == 701
    assert projection["data_rows"] == 700
    assert projection["projection_sha256"] == ("a016ce6370c09c23c8f270e8dabc5330158f2ab4023e77b834bf49fe484a83b5")
    assert projection["normalized_body_sha256"] == ("9a7939360802b4fea8cd15cca93f9eb7b0d2d5fa43b0fe38aa7ff5e66f1fd6e8")


def test_contoso_semantics_match_and_retain_old_pyside6_blank_improvement() -> None:
    projection = _fixture()["contoso_no_image_projection"]
    assert projection["old_pyside6_current_text_equal_after_newline_normalization"] is True
    assert projection["old_pyside6_current_text_sha256"] == (
        "493976dc935bebef208ba1e54a63deeb9b7acd5013fd5d15c5bc72c539ae6274"
    )
    assert projection["all_three_project_semantic_projections_equal"] is True
    assert projection["entries"] == 82
    assert projection["headings"] == 9
    assert projection["table_rows_including_headers"] == 72
    assert projection["empty_sheet_markers"] == 1
    assert projection["projection_sha256"] == ("1f2817bc1dc0aa01bd591f031c2a9711bebd2998867343cdb102708bd25f6c64")
    assert "literal nan" in projection["old_pyside6_improvement_retained"]


def test_contoso_image_option_restores_tk_and_exceeds_old_partial_pyside6() -> None:
    images = _fixture()["contoso_keep_images_projection"]
    assert images["docwen-ref-tk_image_count"] == 6
    assert images["docwen-ref-pyside6_image_count"] == 0
    assert images["docwen-current_image_count"] == 6
    assert images["tk_current_image_bytes_equal"] is True
    assert images["image_size_bytes_each"] == 16528
    assert images["image_sha256_each"] == ("956bc5cb06cad6196955b6d5dc34130bdef297c328e44721f5970c81d266a4f0")
    assert len(images["current_image_names"]) == 6
    assert images["current_image_names"][0] == "PV VTF People Analysis_image1.png"
    assert "restores Tk-primary" in images["classification"]


def test_current_runtime_finalizer_projection_is_complete() -> None:
    runtime = _fixture()["current_runtime_finalizer"]
    financial = runtime["financial_sample"]
    assert financial["artifact_count"] == 1
    assert financial["primary_bytes"] == financial["output_bytes"] == 190323
    assert financial["input_bytes"] == 83418
    assert financial["diagnostics"] == ["SHEET2MD-OK", "FINALIZER_DONE"]
    assert financial["workspace_paths_leaked"] is False

    contoso = runtime["contoso_keep_images"]
    assert contoso["artifact_count"] == 7
    assert contoso["image_artifact_count"] == 6
    assert contoso["primary_bytes"] == 6667
    assert contoso["output_bytes"] == 105835
    assert contoso["input_bytes"] == 7361732
    assert (contoso["sheet_count"], contoso["row_count"], contoso["column_count"]) == (
        9,
        6075,
        26,
    )
    assert (contoso["block_count"], contoso["image_count"]) == (16, 6)
    assert contoso["all_artifacts_under_requested_output_directory"] is True
    assert contoso["diagnostics"] == ["SHEET2MD-OK", "FINALIZER_DONE"]
    assert contoso["workspace_paths_leaked"] is False
