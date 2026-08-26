"""Guards for the VIS-103 real LibreOffice fallback matrix."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest
from tools.validation.source_family import read_source_text

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
FIXTURE_PATH = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_libreoffice_fallback_matrix_semantics.json"
REPORT_NAME = "libreoffice-fallback-real-matrix-2026-07-17.md"


def _read(path: Path) -> str:
    return read_source_text(path)


def _fixture() -> dict[str, Any]:
    return json.loads(_read(FIXTURE_PATH))


def test_libreoffice_fixture_pins_official_isolated_toolchain() -> None:
    data = _fixture()
    toolchain = data["toolchain"]

    assert data["case_id"] == "old_system_libreoffice_fallback_matrix_semantics"
    assert toolchain["version"].startswith("26.2.4.2 ")
    assert toolchain["msi_size_bytes"] == 372539392
    assert toolchain["msi_sha256"] == ("202F26CDA071C5AA4996A5A28412FDDCEB3891DCEB0366982C62650456C0730F")
    isolation = toolchain["isolation"]
    assert isolation["program_on_path"] is True
    assert isolation["bootstrap_user_installation"] == "$ORIGIN/../Data/settings"
    assert isolation["system_winget_registration"] is False
    assert isolation["roaming_profile_created"] is False
    assert isolation["sal_use_vclplugin"] == "svp"
    assert "documentfoundation.org" in toolchain["official_msi_url"]
    assert "Installing_in_parallel/Windows" in toolchain["official_parallel_install_guidance"]


def test_libreoffice_fixture_records_complete_three_project_matrix() -> None:
    data = _fixture()
    matrix = data["matrix"]

    assert data["projects"]["old_tk"]["head"] == "ec9298286cfe1379d5c5470db381577ea43ca0fa"
    assert data["projects"]["old_pyside6"]["head"] == "63db927c5ded920d4994bfede5c7b34c55e2f43e"
    assert data["projects"]["current"]["head_at_probe_start"] == "713e610"
    assert matrix["project_count"] == 3
    assert matrix["case_count_per_project"] == 5
    assert matrix["attempt_count"] == matrix["success_count"] == 15
    assert matrix["all_backends"] == "LibreOffice"
    assert matrix["explicit_empty_com_candidates"] is True
    assert matrix["no_new_lingering_soffice_processes"] is True
    assert len(matrix["shared_semantic_projection_sha256"]) == 64

    cases = {item["case"]: item for item in matrix["cases"]}
    assert set(cases) == {
        "document_odt",
        "document_pdf",
        "spreadsheet_ods",
        "spreadsheet_pdf",
        "presentation_pptx",
    }
    assert cases["document_pdf"]["libreoffice_format"]["current"] == "pdf:writer_pdf_Export"
    assert cases["spreadsheet_pdf"]["libreoffice_format"]["current"] == "pdf:calc_pdf_Export"
    assert cases["spreadsheet_ods"]["semantic_projection"]["formula"] == "of:=SUM([.B3:.B4])"
    assert cases["presentation_pptx"]["semantic_projection"]["shape_counts"] == [2, 2]
    for case in cases.values():
        assert len(case["semantic_sha256"]) == 64
        assert set(case["artifacts"]) == {"old_tk", "old_pyside6", "current"}
        assert all(len(item["sha256"]) == 64 for item in case["artifacts"].values())


def test_libreoffice_fixture_records_physical_and_cancellation_boundaries() -> None:
    data = _fixture()
    physical = data["physical_projection"]
    cancel = data["current_cancellation"]

    assert physical["rendered_page_count"] == 45
    assert physical["dpi"] == 120
    assert physical["cross_project_page_pixels_equal"] is True
    assert {name: item["pages_per_project"] for name, item in physical["surfaces"].items()} == {
        "document_pdf": 6,
        "spreadsheet_pdf": 1,
        "document_odt_render": 6,
        "presentation_pptx_render": 2,
    }
    assert all(len(item["ordered_page_hash_sha256"]) == 64 for item in physical["surfaces"].values())
    formula = physical["shared_formula_boundary"]
    assert formula["glyph"] == "U+2751"
    assert formula["occurrences"] == {"old_tk": 1, "old_pyside6": 1, "current": 1}
    assert "not a current-only regression" in formula["classification"]

    assert cancel == {
        "cancel_after_seconds": 0.5,
        "elapsed_seconds": 0.61,
        "success": False,
        "message": "cancelled",
        "output_exists": False,
        "output_directory_entry_count": 0,
        "new_lingering_process_count": 0,
        "comparison_boundary": (
            "Current proves mid-process cancellation. The reference implementations only prove pre-call "
            "cancellation by source trace and are not claimed equivalent."
        ),
    }


def test_libreoffice_filtered_output_repair_is_guarded_at_owner_and_consumer() -> None:
    data = _fixture()["filtered_pdf_repair"]
    bridge = _read(PROJECT_ROOT / "packages" / "core" / "src" / "docwen_core" / "office_bridge.py")
    bridge_tests = _read(PROJECT_ROOT / "packages" / "core" / "tests" / "test_office_bridge_*.py")
    print_converter = _read(
        PROJECT_ROOT
        / "packages"
        / "plugins"
        / "print"
        / "src"
        / "docwen_plugin_print"
        / "paged_output"
        / "converter.py"
    )

    assert data["reproduced_before_fix"] is True
    assert data["format"] == "pdf:writer_pdf_Export"
    assert data["libreoffice_created_size_bytes"] == 32664
    assert 'convert_to.partition(":")[0].strip().lstrip(".")' in bridge
    assert "test_libreoffice_conversion_uses_extension_before_filter_name" in bridge_tests
    assert "pdf:writer_pdf_Export" in print_converter
    assert "pdf:calc_pdf_Export" in print_converter


def test_libreoffice_fixture_pins_authoritative_external_subset() -> None:
    evidence = _fixture()["external_evidence"]

    assert evidence["root"] == "D:/docwen-parity/vis103-libreoffice-26.2.4-713e610-v1"
    assert evidence["root_file_count_including_diagnostics"] == 89
    assert evidence["root_size_bytes_including_diagnostics"] == 7630840
    assert evidence["authoritative_subset_file_count"] == 30
    assert evidence["authoritative_subset_size_bytes"] == 3165756
    assert len(evidence["contact_sheets"]) == 3
    assert all(len(item["sha256"]) == 64 for item in evidence["contact_sheets"].values())
    assert "excluded from the authoritative acceptance subset" in evidence["diagnostic_boundary"]
