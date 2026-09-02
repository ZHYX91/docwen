from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract
ROOT = Path(__file__).resolve().parents[2]
FIXTURE = ROOT / "tests/fixtures/golden/old_system_apache_poi_ooxml_signature_semantics.json"


def _data() -> dict:
    return json.loads(FIXTURE.read_text(encoding="utf-8"))


def test_official_sources_are_frozen_signed_unsigned_and_not_distributed() -> None:
    data = _data()
    repo = data["source_repository"]
    assert (repo["owner"], repo["repository"], repo["commit"], repo["tree"]) == (
        "apache",
        "poi",
        "86e967d9b28d6a322a87ae8fcbf2a7eeb56cef96",
        "712aa87cce1ca0416313b0f604372cd92b6befa6",
    )
    assert (repo["tree_entries"], repo["tree_truncated"], repo["path"]) == (
        6405,
        False,
        "test-data/xmldsign",
    )
    sources = data["sources"]
    assert len(sources) == 10
    assert sum(item["signature_xml_count"] for item in sources.values()) == 8
    assert sources["hello-world-signed-twice.docx"]["signature_xml_count"] == 2
    for name in sources:
        assert not any((ROOT / "tests/fixtures/files").rglob(name))


def test_office_recognition_is_not_misreported_as_current_trust_validation() -> None:
    facts = _data()["signature_source_facts"]
    assert facts["office_16_signature_counts"] == {
        "ordinary_signed": 1,
        "double_signed_docx": 2,
        "unsigned": 0,
    }
    assert facts["office_reports_all_historical_signatures_valid_now"] is False
    assert facts["office_reports_all_historical_certificates_expired"] is True
    assert "not a currently valid or trusted" in facts["trust_boundary"]
    contract = _data()["public_contract"]
    assert contract["all_three_accept_docx_xlsx_pptx"] is True
    assert contract["all_three_expose_signature_validation"] is False
    assert contract["all_three_expose_signature_preservation"] is False
    assert contract["pptx_to_pdf_public_route"] is False


def test_markdown_and_pdf_matrices_record_parity_without_raw_hash_overclaim() -> None:
    data = _data()
    markdown = data["markdown_matrix"]
    assert (markdown["production_conversions"], markdown["successes"]) == (30, 30)
    assert markdown["tk_old_raw_byte_equal"] == 10
    assert markdown["cross_project_normalized_body_equal"] == 10
    assert markdown["source_signature_marker_exposed"] is False
    pdf = data["pdf_matrix"]
    assert (pdf["production_conversions"], pdf["successes"]) == (21, 21)
    assert pdf["docx_cross_project_exact_page_text_geometry_pixel_equal"] == 4
    assert pdf["xlsx_ms_office_cross_project_exact_page_text_geometry_pixel_equal"] is True
    assert pdf["xlsx_hello_world_max_origin_delta_points"] == {"x": 0.25, "y": 0.1}
    assert pdf["current_repeat_same_xlsx_pixel_projection_equal"] == 4
    assert pdf["raw_pdf_sha_is_nondeterministic"] is True
    assert pdf["pdf_signature_fields_present"] is False
    assert pdf["final_new_wps_processes"] == 0


def test_shared_signature_security_ux_boundary_remains_unaccepted() -> None:
    classification = _data()["classification"]
    assert classification["current_only_functional_regression_found"] is False
    assert classification["production_change_made"] is False
    assert classification["focused_source_content_conversion_evidence"] == "strong"
    assert classification["signature_validation_preservation_evidence"] == ("explicitly unsupported shared boundary")
    assert len(classification["known_blockers_not_accepted"]) == 4
    assert "remain open" in classification["completion_effect"]
