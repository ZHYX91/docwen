"""Fail-closed evidence guard for the selected POLICY-03=A implementation."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]
GOLDEN = ROOT / "tests/fixtures/golden"
FIXTURE = GOLDEN / "current_policy03_preserved_presentation_payloads_semantics.json"
CONVERTER = ROOT / "packages/plugins/presentation/src/docwen_plugin_presentation/pptx_md/converter.py"
EVALUATOR = ROOT / "tools/validation/evaluate_policy03_preserved_payloads.py"


def _fixture() -> dict[str, object]:
    return json.loads(FIXTURE.read_text(encoding="utf-8"))


def test_policy03_frozen_sources_and_exact_payload_oracle_are_pinned() -> None:
    fixture = _fixture()
    assert fixture["evidence_id"] == "VIS-2026-07-23-204"
    assert fixture["policy"] == "POLICY-03=A"
    assert fixture["classification"] == "FIXED_AND_VERIFIED"
    evidence = fixture["authoritative_current_evidence"]
    assert (
        evidence["file_count"],
        evidence["total_bytes"],
        evidence["result_json_bytes"],
    ) == (20, 825_745, 22_943)
    assert evidence["result_json_sha256"] == ("6e21a7fa6f438e3a0b4f8ffc47f6206d89b04b649ff0acd087b4800a4e5907d4")
    assert evidence["passed"] is True

    chart = fixture["sources"]["bar-chart.pptx"]["current"]
    assert chart["ordered_rows"] == [
        ["1st Qtr", "8.200000000000001"],
        ["2nd Qtr", "3.2"],
        ["3rd Qtr", "1.4"],
        ["4th Qtr", "1.2"],
    ]
    assert chart["embedded_workbook"]["exact"] is True
    assert chart["embedded_workbook"]["linked"] is True
    assert chart["snapshot_warning"] == {
        "code": "PPTX-CHART-SNAPSHOT-UNAVAILABLE",
        "count": 1,
        "location": "slide 1: chart 1",
    }
    assert chart["snapshot_artifact_count"] == 0

    for source, payloads in (
        ("EmbeddedAudio.pptx", ("audio", "poster")),
        ("EmbeddedVideo.pptx", ("video", "poster")),
    ):
        current = fixture["sources"][source]["current"]
        for payload in payloads:
            assert current[payload]["exact"] is True
            assert current[payload]["linked"] is True


def test_policy03_same_basename_and_implementation_shape_remain_closed() -> None:
    fixture = _fixture()
    same = fixture["same_basename"]
    assert same == {
        "conversion_count": 3,
        "primary_names": ["same.md", "same_001.md", "same_002.md"],
        "primary_names_unique": True,
        "all_payload_targets_reachable": True,
        "payloads_exact": True,
    }
    boundaries = fixture["boundaries"]
    assert boundaries["source_immutable"] is True
    assert boundaries["new_office_processes"] == []
    assert boundaries["accepted_payload_omission"] is False
    assert boundaries["generated_binaries_checked_in"] is False

    source = CONVERTER.read_text(encoding="utf-8")
    for token in (
        "_preserve_chart_payload",
        "_preserve_media_payload",
        "PPTX-CHART-SNAPSHOT-UNAVAILABLE",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "audio/mpeg",
        "video/mp4",
    ):
        assert token in source
    assert EVALUATOR.is_file()
