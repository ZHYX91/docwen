"""Contracts for VIS-022 SmartDoc rich RTF-to-ODT artifact evidence."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

PROJECT_ROOT = Path(__file__).resolve().parents[2]
FIXTURE_NAME = "old_system_smartdoc_rtf_odt_two_hop_semantics.json"
FIXTURE_PATH = PROJECT_ROOT / "tests" / "fixtures" / "golden" / FIXTURE_NAME
REPORT_NAME = "smartdoc-rtf-odt-two-hop-artifact-2026-07-14.md"


def _fixture() -> dict[str, object]:
    return json.loads(FIXTURE_PATH.read_text(encoding="utf-8"))


def test_rtf_odt_fixture_records_three_project_semantic_and_package_projection() -> None:
    fixture = _fixture()

    assert fixture["golden_id"] == "GOLDEN-009"
    assert fixture["classification"] == ("focused_real_external_artifact_evidence_not_overall_smartdoc_pass")
    assert fixture["input_document"]["size_bytes"] == 62215
    assert fixture["input_document"]["sha256"] == ("7548f563c27065434cd75770532cbf89ec4459863e9ba5a6535ec4cce4780017")
    assert fixture["input_document"]["generation_backend"] == "WPS Writer"

    route = fixture["route"]
    assert (route["source_format"], route["hub_format"], route["target_format"]) == (
        "rtf",
        "docx",
        "odt",
    )
    assert route["candidate_policy"]["first_leg"] == [
        "WPS Writer",
        "Microsoft Word",
        "LibreOffice",
    ]
    assert route["candidate_policy"]["second_leg"] == ["Microsoft Word", "LibreOffice"]
    assert route["current_observed_backend"] == "WPS Writer -> Microsoft Word"

    projection = fixture["normalized_common_projection"]
    assert projection["section_landscape"] == [False, True]
    assert projection["comments"] == ["DOCWEN RTF ODT COMMENT 2026"]
    assert projection["field_count"] == 1
    assert projection["media_count"] == 1
    assert projection["media_size_bytes"] == 2006
    assert projection["media_sha256"] == ("acff2b06fddc6f2f6c6d4723f8890dbcc2776dbdd16968e485abaf8388f82b26")
    assert projection["table"] == [
        ["Item", "Qty", "Note"],
        ["RtfOrange", "22", "TABLE-RTF-ODT"],
    ]
    assert {item["text"]: (item["bold"], item["italic"]) for item in projection["styled_runs"]} == {
        "BOLD-RTF-ODT": (True, False),
        "ITALIC-RTF-ODT": (False, True),
    }

    package = fixture["odt_package_projection"]
    assert package["is_zip"] is True
    assert package["mimetype"] == "application/vnd.oasis.opendocument.text"
    assert package["annotation_count"] == 1
    assert package["draw_image_count"] == 1
    assert package["embedded_media_path"] == "media/image1.png"
    assert package["tracked_changes_count"] == 0
    assert package["content_xml_strict_parse"] is False
    assert "numeric xml:id" in package["content_xml_strict_parse_boundary"]

    projects = fixture["projects"]
    assert set(projects) == {"docwen-ref-tk", "docwen-ref-pyside6", "docwen-current"}
    for project in projects.values():
        assert project["projection_matches_common"] is True
        assert project["odt_package_matches_common"] is True
        assert project["size_bytes"] > 0
        assert len(project["sha256"]) == 64
        assert project["roundtrip_docx_size_bytes"] > 0
        assert len(project["roundtrip_docx_sha256"]) == 64

    current = projects["docwen-current"]
    assert current["filename"] == "smartdoc-rich-rtf-input.odt"
    assert current["artifact"] == {
        "suggested_name": "smartdoc-rich-rtf-input.odt",
        "source_format": "rtf",
        "target_format": "odt",
        "backend": "WPS Writer -> Microsoft Word",
    }
    assert current["metrics"] == {
        "input_bytes": 62215,
        "output_bytes": 12490,
        "backend": "WPS Writer -> Microsoft Word",
    }
    assert fixture["naming_decision"]["classification"] == ("current_architecture_improvement_not_a_parity_gap")
    assert fixture["process_governance"]["new_relevant_processes_after_conversion_and_readback"] == []
