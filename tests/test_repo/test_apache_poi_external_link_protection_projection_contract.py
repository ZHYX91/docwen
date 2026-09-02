from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract
ROOT = Path(__file__).resolve().parents[2]
FIXTURE = ROOT / "tests/fixtures/golden/old_system_apache_poi_external_link_protection_semantics.json"


def _data() -> dict:
    return json.loads(FIXTURE.read_text(encoding="utf-8"))


def test_sources_are_frozen_feature_screened_and_not_distributed() -> None:
    data = _data()
    repo = data["source_repository"]
    assert (repo["owner"], repo["repository"], repo["commit"], repo["tree"]) == (
        "apache",
        "poi",
        "86e967d9b28d6a322a87ae8fcbf2a7eeb56cef96",
        "712aa87cce1ca0416313b0f604372cd92b6befa6",
    )
    assert (repo["tree_entries"], repo["tree_truncated"]) == (6405, False)
    sources = data["sources"]
    assert sources["45431.xlsm"]["vba_part"] == "xl/vbaProject.bin"
    assert sources["link-external-workbook-b.xlsx"]["cached_value"] == 30.0
    assert len(sources) == 6
    for name in sources:
        assert not any((ROOT / "tests/fixtures/files").rglob(name))


def test_macro_format_is_not_misreported_as_public_parity_support() -> None:
    macro = _data()["macro_contract"]
    assert macro["sample_has_vba_project"] is True
    assert macro["tk_supported"] is False
    assert macro["old_pyside6_supported"] is False
    assert macro["current_core_supported"] is False
    assert macro["current_gui_supported"] is False
    assert (macro["current_cli_dry_run_error"], macro["format"]) == ("invalid_input", "xlsm")
    assert "unreachable residue" in macro["classification"]


def test_real_ods_matrix_records_current_cleanup_improvement_without_overclaim() -> None:
    data = _data()
    matrix = data["ods_execution_matrix"]
    assert matrix["tk"]["external_link"] == matrix["current"]["external_link"] == "success"
    assert matrix["tk"]["all_locked_after_two_failures"] == "failed"
    assert matrix["tk"]["all_locked_isolated"] == matrix["current"]["all_locked"] == "success"
    assert matrix["old_pyside6"]["plugin_route"] == "NOT_IMPLEMENTED"
    assert matrix["old_pyside6"]["legacy_external_link_probe"] == "hung_without_result"
    assert matrix["old_pyside6"]["final_excel_processes"] == 0
    assert matrix["current"]["final_excel_processes"] == 0
    assert matrix["current"]["workbook_password"] == {
        "success": False,
        "error_code": "dependency_missing",
        "exit_code": 4,
    }
    assert data["classification"]["current_only_functional_regression_found"] is False
    assert data["classification"]["production_change_made"] is False


def test_equal_outputs_keep_external_link_and_password_boundaries_open() -> None:
    data = _data()
    semantic = data["semantic_projection"]
    assert semantic["tk_current_external_equal"] is True
    assert semantic["source_cached_value"] == 30.0
    assert semantic["output_cached_value"] == "#REF!"
    assert semantic["relative_target_rewritten_to_absolute_file_uri"] is True
    assert semantic["tk_current_all_locked_equal"] is True
    assert semantic["all_locked_protection_preserved"] is True
    diagnostics = data["direct_excel_diagnostics"]
    assert diagnostics["save_as_ods_hresult"] == -2146827284
    assert diagnostics["update_links_3_external_value"] == 30.0
    assert diagnostics["update_links_3_embedded_external_values"] == [10.0, 20.0]
    assert "not safe" in diagnostics["security_boundary"]
    assert len(data["classification"]["known_blockers_not_accepted"]) == 4
