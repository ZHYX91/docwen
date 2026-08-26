"""Fail-closed guards for the VIS-2026-07-17-122 physical-oracle boundary."""

from __future__ import annotations

import json
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]
FIXTURE = ROOT / "tests/fixtures/golden/old_system_docx_comprehensive_roundtrip_semantics.json"
SOURCE = ROOT / "tests/fixtures/golden/md_to_docx_old/sample_golden.docx"
REPORT_NAME = "historical-comprehensive-docx-physical-oracle-boundary-2026-07-17.md"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _fixture() -> dict[str, object]:
    return json.loads(FIXTURE.read_text(encoding="utf-8"))


def _boundary() -> dict[str, object]:
    return _fixture()["physical_oracle_boundary"]


def test_historical_source_identity_and_nonpositive_geometry_are_exact() -> None:
    fixture = _fixture()
    assert fixture["source"]["bytes"] == 39_146
    assert fixture["source"]["sha256"] == ("7903834d9cadddf48287695e3ab1b2f65b5cc052bbabdb92a97e8be139eadd53")

    with zipfile.ZipFile(SOURCE) as archive:
        document = ET.fromstring(archive.read("word/document.xml"))
    section = document.find(f".//{W}sectPr")
    assert section is not None
    page = section.find(f"{W}pgSz")
    margins = section.find(f"{W}pgMar")
    assert page is not None and margins is not None
    assert (int(page.attrib[f"{W}w"]), int(page.attrib[f"{W}h"])) == (1875, 2652)
    assert [int(margins.attrib[f"{W}{name}"]) for name in ("top", "right", "bottom", "left")] == [
        1440,
        1440,
        1440,
        1440,
    ]

    geometry = _boundary()["source_geometry"]
    assert geometry["usable_content_twips"] == {"width": -1005, "height": -228}
    assert geometry["usable_content_area_is_nonpositive"] is True


def test_external_projection_and_three_project_result_are_pinned() -> None:
    boundary = _boundary()
    assert boundary["evidence_id"] == "VIS-2026-07-17-122"
    assert boundary["source_word_projection"] == {
        "pages": 1516,
        "pdf_page_points": {"width": 93.72, "height": 132.6},
        "nonempty_pdf_pages": 1159,
        "paragraph_count": 110,
        "table_count": 2,
        "field_count": 1,
        "footnote_count": 2,
        "endnote_count": 2,
        "hyperlink_count": 1,
        "omath_count": 3,
    }
    execution = boundary["three_project_execution"]
    assert execution["production_conversions"] == 9
    assert execution["all_conversions_successful"] is True
    assert execution["same_word_object_projection_per_target"] is True
    assert (execution["doc_pages_per_project"], execution["doc_repaginate_and_pdf_export_failures"]) == (601, 3)
    assert (execution["rtf_pages_per_project"], execution["rtf_pdf_exports"]) == (1513, 3)
    assert (execution["odt_pages_per_project"], execution["odt_pdf_exports"]) == (1324, 3)


def test_source_role_and_nonclosure_classification_fail_closed() -> None:
    classification = _boundary()["classification"]
    assert classification == {
        "historical_source_is_valid_semantic_fixture": True,
        "historical_source_is_valid_physical_or_pagination_oracle": False,
        "current_only_conversion_or_object_regression_found": False,
        "production_change_made": False,
        "vis101_fresh_a4_output_evidence_remains_valid": True,
        "broad_doc_docx_rtf_odt_physical_parity_closed": False,
        "overall_parity_closed": False,
    }
    cleanup = _boundary()["process_cleanup"]
    assert cleanup["all_isolated_probes_cleaned_within_30_seconds"] is True
    assert cleanup["maximum_observed_cleanup_seconds"] == 2
    assert cleanup["termination_command_used"] is False
