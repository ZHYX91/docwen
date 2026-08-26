"""Contracts for VIS-020 rich SmartDoc outbound artifact evidence."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

PROJECT_ROOT = Path(__file__).resolve().parents[2]
FIXTURE_PATH = (
    PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_smartdoc_rich_outbound_fidelity_semantics.json"
)
REPORT_NAME = "smartdoc-rich-outbound-fidelity-2026-07-14.md"


def _fixture() -> dict[str, object]:
    return json.loads(FIXTURE_PATH.read_text(encoding="utf-8"))


def test_rich_smartdoc_fixture_records_three_project_normalized_projection() -> None:
    fixture = _fixture()

    assert fixture["golden_id"] == "GOLDEN-009"
    assert fixture["classification"] == ("focused_real_external_artifact_evidence_not_overall_smartdoc_pass")
    input_document = fixture["input_document"]
    assert input_document["size_bytes"] == 42591
    assert input_document["sha256"] == ("7f17662cca68b1668bbfd1c46bb6cddf87a4c0fd17bc92d0a8737472976c8a0f")

    routes = fixture["routes"]
    assert routes["doc"]["candidate_chain"] == [
        "WPS Writer",
        "Microsoft Word",
        "LibreOffice",
    ]
    assert routes["rtf"]["candidate_chain"] == routes["doc"]["candidate_chain"]
    assert routes["odt"]["candidate_chain"] == ["Microsoft Word", "LibreOffice"]
    assert routes["doc"]["signature_hex_prefix"] == "d0cf11e0a1b11ae1"
    assert routes["rtf"]["signature_hex_prefix"] == "7b5c72746631"
    assert routes["odt"]["is_zip"] is True

    projection = fixture["normalized_common_projection"]
    assert projection["deleted_marker_visible"] is False
    assert projection["section_landscape"] == [False, True]
    assert projection["hyperlink_count"] == 1
    assert projection["media_count"] == 1
    assert projection["media_sha256"] == ("cea0e700deb7e3b8e3934924520c806944d5da7755bb17335c0e5b42de7a9a81")
    assert projection["comments"] == ["DOCWEN COMMENT BODY 2026"]
    assert projection["table"] == [
        ["Item", "Qty", "Note"],
        ["RichApple", "11", "TABLE-MARKER"],
    ]
    assert {item["text"]: (item["bold"], item["italic"]) for item in projection["styled_runs"]} == {
        "BOLD-MARKER": (True, False),
        "ITALIC-MARKER": (False, True),
    }

    boundaries = fixture["format_boundaries"]
    for target in ("doc", "rtf"):
        assert boundaries[target]["revision_counts"] == {"insertions": 1, "deletions": 1}
        assert boundaries[target]["roundtrip_media_count"] == 1
    assert boundaries["odt"]["revision_counts"] == {"insertions": 0, "deletions": 0}
    assert boundaries["odt"]["direct_annotation_count"] == 1
    assert boundaries["odt"]["direct_draw_image_count"] == 1
    assert boundaries["odt"]["content_xml_strict_parse"] is False
    assert "not an XML NCName" in boundaries["odt"]["content_xml_strict_parse_boundary"]

    projects = fixture["projects"]
    assert set(projects) == {"docwen-ref-tk", "docwen-ref-pyside6", "docwen-current"}
    for project in projects.values():
        assert set(project["artifacts"]) == {"doc", "rtf", "odt"}
        for artifact in project["artifacts"].values():
            assert artifact["size_bytes"] > 0
            assert len(artifact["sha256"]) == 64
            assert artifact["roundtrip_docx_size_bytes"] > 0
            assert len(artifact["roundtrip_docx_sha256"]) == 64

    current = projects["docwen-current"]
    assert current["cli_duration_seconds"] == {"doc": 1.877, "rtf": 1.808, "odt": 6.539}
    assert fixture["process_governance"]["new_relevant_processes_after_conversion_and_readback"] == []
    assert "remains open" in fixture["environment_boundary"]["libreoffice"]
