"""Fail-closed contracts for VIS-2026-07-17-129 Presentation media evidence."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]
GOLDEN = ROOT / "tests" / "fixtures" / "golden"
FIXTURE = GOLDEN / "old_system_pptx_to_markdown_semantics.json"
REPORT_NAME = "apache-poi-presentation-chart-audio-video-matrix-2026-07-17.md"


def _addendum() -> dict[str, object]:
    fixture = json.loads(FIXTURE.read_text(encoding="utf-8"))
    return fixture["chart_audio_video_addendum"]


def test_official_source_identities_and_package_payloads_are_pinned() -> None:
    addendum = _addendum()
    assert addendum["evidence_id"] == "VIS-2026-07-17-129"
    assert addendum["source_repository"] == "apache/poi"
    assert addendum["source_commit"] == "86e967d9b28d6a322a87ae8fcbf2a7eeb56cef96"
    sources = addendum["sources"]

    assert sources["bar-chart.pptx"]["git_blob"] == ("e4d2613046ab69e2d0a5c529b41cbcaa49ac4e30")
    assert (sources["bar-chart.pptx"]["bytes"], sources["bar-chart.pptx"]["sha256"]) == (
        44_410,
        "79e1d218bfb2903e8dc8425a6b1997d9c1976f5a5f025bada85b0c47b5777969",
    )
    assert sources["bar-chart.pptx"]["chart_xml"]["sha256"] == (
        "c6f133f62ccab33ad855c6399408439e3e7d780ec9191f01a421a909259bb4d3"
    )
    assert sources["bar-chart.pptx"]["embedded_workbook"]["sha256"] == (
        "89673f803b955c3f553900ddfd406a80babe5a543e8f8401bfa7b2b7834cae22"
    )

    assert sources["EmbeddedAudio.pptx"]["git_blob"] == ("ab12d00acbea0f74607b8fc9f077f26a2d1953ea")
    assert sources["EmbeddedAudio.pptx"]["poster"]["sha256"] == (
        "b0151c2c2e3cf64bc37a7bb9d8b8b98d4c4fccf7b6af4c08c4f847a79f9db0da"
    )
    assert sources["EmbeddedAudio.pptx"]["audio"]["sha256"] == (
        "0244590f2b4bcb62352b574e78bea940e8d89cfa69823b5208ef4c43e0abcb44"
    )

    assert sources["EmbeddedVideo.pptx"]["git_blob"] == ("f7954228a8d95e05bcaa50364b220e9780c68e92")
    assert sources["EmbeddedVideo.pptx"]["poster"]["sha256"] == (
        "f5516c6cae484df63ce03db77fb69b778660916b9207de5a4e04aa5e3b72908d"
    )
    assert sources["EmbeddedVideo.pptx"]["video"]["sha256"] == (
        "21c3b5d779abe3bc2ee886a6d2455202800537fa31fe367d11563da16cbf8040"
    )


def test_nine_conversions_and_normalized_three_project_projections_are_exact() -> None:
    addendum = _addendum()
    execution = addendum["execution"]
    assert execution["public_route"] == "pptx->md"
    assert execution["success_count"] == execution["expected_success_count"] == 9
    assert execution["conversion_process_added"] == []
    assert execution["conversion_process_removed"] == []
    assert execution["termination_command_used"] is False

    outputs = addendum["outputs"]
    for source in ("bar-chart.pptx", "EmbeddedAudio.pptx", "EmbeddedVideo.pptx"):
        projection = outputs[source]
        assert projection["old_markdown_raw_exact"] is True
        assert projection["three_project_semantic_exact_after_resource_hash_normalization"] is True
        assert projection["three_project_resource_bytes_exact"] is True

    assert outputs["bar-chart.pptx"]["three_project_raw_markdown_exact"] is True
    assert outputs["EmbeddedAudio.pptx"]["three_project_raw_markdown_exact"] is False
    assert outputs["EmbeddedAudio.pptx"]["resource_count_each"] == [1, 1, 1]
    assert outputs["EmbeddedVideo.pptx"]["three_project_raw_markdown_exact"] is True
    assert outputs["EmbeddedVideo.pptx"]["resource_count_each"] == [0, 0, 0]

    runtime = addendum["current_runtime"]
    assert runtime["diagnostic_codes_each"] == ["PPTX2MD-OK", "FINALIZER_DONE"]
    assert (
        runtime["bar_chart_artifact_count"],
        runtime["audio_artifact_count"],
        runtime["video_artifact_count"],
    ) == (1, 2, 1)


def test_chart_oracle_and_shared_semantic_omission_remain_explicit() -> None:
    addendum = _addendum()
    chart = addendum["source_fidelity"]["chart"]
    assert (chart["powerpoint_shape_type"], chart["chart_type"], chart["series_count"]) == (
        3,
        57,
        1,
    )
    assert chart["chart_title"] == "Sales"
    assert chart["visible_tokens_absent_from_all_outputs"] == [
        "Sales",
        "1st Qtr",
        "2nd Qtr",
        "3rd Qtr",
        "4th Qtr",
        "0",
        "2",
        "4",
        "6",
        "8",
        "10",
    ]
    assert chart["source_slide_title_retained_by_all_projects"] is True
    assert chart["output_chart_snapshot_count_each"] == [0, 0, 0]

    boundary = addendum["shared_unaccepted_boundary"]
    assert boundary["classification"] == "known_blocker_not_accepted_difference"
    assert "omit visible chart series/category/axis/title semantics" in boundary["detail"]
    assert "Equal omission is not source-faithful conversion" in boundary["detail"]
    assert "chart and embedded-media extraction/preservation policies" in boundary["acceptance_condition"]
