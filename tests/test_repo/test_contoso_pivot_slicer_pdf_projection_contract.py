from __future__ import annotations

import json
import tomllib
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

PROJECT_ROOT = Path(__file__).resolve().parents[2]
FIXTURE = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_real_financial_workbook_batch_semantics.json"
PRINT_CONVERTER = (
    PROJECT_ROOT / "packages" / "plugins" / "print" / "src" / "docwen_plugin_print" / "paged_output" / "converter.py"
)


def _addendum() -> dict:
    fixture = json.loads(FIXTURE.read_text(encoding="utf-8"))
    return fixture["contoso_pivot_slicer_pdf_addendum"]


def test_contoso_source_has_real_pivot_slicer_and_data_model_parts() -> None:
    profile = _addendum()["source_package_profile"]
    assert profile["package_entries"] == 260
    assert profile["sheet_count"] == 9
    assert profile["pivot_table_parts"] == 3
    assert profile["pivot_cache_definition_parts"] == 4
    assert profile["slicer_cache_parts"] == 3
    assert profile["slicer_parts"] == 1
    assert profile["slicer_defined_names"] == [
        "Slicer_Fiscal_Year",
        "Slicer_Month",
        "Slicer_Sub_Class",
    ]
    assert "ThisWorkbookDataModel" in profile["connection_names"]
    assert profile["model_item_bytes"] == 6094848
    assert profile["model_item_sha256"] == ("1ECE8FB5A7350EE45677452BE5F1461C3DC7563C489B6CC852AA945CA956ED0D")


def test_print_priority_consumer_matches_authoritative_base_config() -> None:
    from docwen_plugin_print.paged_output.converter import _DEFAULT_PRIORITIES

    config = tomllib.loads((PROJECT_ROOT / "configs" / "software.toml").read_text(encoding="utf-8"))
    special = config["special_conversions"]
    assert list(_DEFAULT_PRIORITIES["document"]) == special["document_to_pdf"]
    assert list(_DEFAULT_PRIORITIES["spreadsheet"]) == special["spreadsheet_to_pdf"]

    source = PRINT_CONVERTER.read_text(encoding="utf-8")
    assert 'f"software.special_conversions.{category}_to_pdf"' in source
    assert "_convert_with_configured_priority" in source
    assert "context.config.get" in source

    fix = _addendum()["current_priority_consumer_fix"]
    assert fix["initial_regression_result"] == "3 failed before production repair"
    assert fix["post_fix_print_package_result"] == "19 passed"
    assert fix["default_spreadsheet_order"] == special["spreadsheet_to_pdf"]
    assert fix["default_document_order"] == special["document_to_pdf"]


def test_three_backend_physical_matrix_records_strong_and_limited_results() -> None:
    addendum = _addendum()
    matrix = addendum["pdf_matrix"]
    assert matrix["excel"]["all_three_ordered_page_pixels_equal"] is True
    assert matrix["excel"]["page_2_slicer_anchors"] == [
        "Fiscal Year",
        "Month",
        "Sub Class",
    ]
    assert matrix["excel"]["power_view_message_pages"] == [4, 5, 6, 8, 9, 10]
    assert matrix["wps"]["tk_current_ordered_page_pixels_equal"] is True
    assert matrix["wps"]["all_three_ordered_page_pixels_equal"] is False
    assert matrix["wps"]["page_2_slicer_anchors_present"] is False
    assert matrix["libreoffice"]["all_three_ordered_page_pixels_equal"] is True
    assert matrix["libreoffice"]["page_2_slicer_anchors_present"] is False
    for backend in matrix.values():
        assert backend["all_three_successful"] is True
        assert backend["page_count_each"] == 10
        assert backend["all_three_pypdf_text_equal"] is True
        assert backend["all_three_pdfplumber_text_equal"] is True

    classification = addendum["classification"]
    assert classification["current_only_config_consumer_gap_found_and_fixed"] is True
    assert classification["tk_primary_default_wps_physical_parity"] is True
    assert classification["excel_high_fidelity_option_parity"] is True
    assert classification["broad_editable_pivot_powerpivot_slicer_fidelity_closed"] is False
    assert classification["overall_parity_closed"] is False
