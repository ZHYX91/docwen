"""Contracts for VIS-021 SmartDoc ODT-to-RTF two-hop artifact evidence."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

PROJECT_ROOT = Path(__file__).resolve().parents[2]
FIXTURE_NAME = "old_system_smartdoc_odt_rtf_two_hop_semantics.json"
FIXTURE_PATH = PROJECT_ROOT / "tests" / "fixtures" / "golden" / FIXTURE_NAME
REPORT_NAME = "smartdoc-odt-rtf-two-hop-artifact-2026-07-14.md"


def _fixture() -> dict[str, object]:
    return json.loads(FIXTURE_PATH.read_text(encoding="utf-8"))


def test_two_hop_fixture_records_three_project_artifact_projection() -> None:
    fixture = _fixture()

    assert fixture["golden_id"] == "GOLDEN-009"
    assert fixture["classification"] == ("focused_real_external_artifact_evidence_not_overall_smartdoc_pass")
    assert fixture["input_document"]["size_bytes"] == 12012
    assert fixture["input_document"]["sha256"] == ("188ed73f96f40dda6def98a48df3c36ed0945decc3778315214e2aaa60fc3568")

    route = fixture["route"]
    assert (route["source_format"], route["hub_format"], route["target_format"]) == (
        "odt",
        "docx",
        "rtf",
    )
    assert route["old_candidate_policy"]["first_leg"] == ["Microsoft Word", "LibreOffice"]
    assert route["old_candidate_policy"]["second_leg"] == [
        "WPS Writer",
        "Microsoft Word",
        "LibreOffice",
    ]
    assert route["current_observed_backend"] == "Microsoft Word -> WPS Writer"
    assert route["signature_hex_prefix"] == "7b5c72746631"

    projection = fixture["normalized_common_projection"]
    assert projection["section_landscape"] == [False, True]
    assert projection["comments"] == ["DOCWEN TWO-HOP COMMENT 2026"]
    assert projection["media_count"] == 1
    assert projection["field_display_marker_retained"] is True
    assert projection["field_container_after_odt_normalization"] == 0
    assert projection["table"] == [
        ["Item", "Qty", "Note"],
        ["TwoHopApple", "21", "TABLE-TWO-HOP"],
    ]
    assert {item["text"]: (item["bold"], item["italic"]) for item in projection["styled_runs"]} == {
        "BOLD-TWO-HOP": (True, False),
        "ITALIC-TWO-HOP": (False, True),
    }

    projects = fixture["projects"]
    assert set(projects) == {"docwen-ref-tk", "docwen-ref-pyside6", "docwen-current"}
    for project in projects.values():
        assert project["projection_matches_common"] is True
        assert project["size_bytes"] > 0
        assert len(project["sha256"]) == 64
        assert project["roundtrip_docx_size_bytes"] > 0
        assert len(project["roundtrip_docx_sha256"]) == 64

    current = projects["docwen-current"]
    assert current["filename"] == "smartdoc-two-hop-rich.rtf"
    assert current["artifact"] == {
        "suggested_name": "smartdoc-two-hop-rich.rtf",
        "source_format": "odt",
        "target_format": "rtf",
        "backend": "Microsoft Word -> WPS Writer",
    }
    assert current["metrics"] == {
        "input_bytes": 12012,
        "output_bytes": 70146,
        "backend": "Microsoft Word -> WPS Writer",
    }
    assert fixture["naming_decision"]["classification"] == ("current_architecture_improvement_not_a_parity_gap")
    assert fixture["process_governance"]["new_relevant_processes_after_conversion_and_readback"] == []
