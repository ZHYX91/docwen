from __future__ import annotations

import pytest
from tools.windows_shards.merge_reports import _visibility

pytestmark = pytest.mark.unit


def test_reports_preserve_slow_calls_and_legal_repeated_subprocesses() -> None:
    def report(items):
        return {
            "generated_at": "now",
            "pytest": {"selection": "full", "xdist": {"enabled": True, "numprocesses": 4, "dist": "loadfile"}},
            "items": items,
            "summary": {"count": len(items), "timeout_count": 1},
        }

    a = {"command": "docwen", "duration_seconds": 1.0, "timed_out": True}
    b = {"command": "docwen", "duration_seconds": 3.0, "timed_out": False}
    left, right = report([a]), report([a, b])
    right["pytest"]["xdist"]["numprocesses"] = 2
    merged = _visibility("subprocess_report.json", left, right)
    assert merged["pytest"]["xdist"]["numprocesses"] == 6
    assert merged["items"] == [b, a, a]
    assert merged["summary"] == {"count": 3, "timeout_count": 2}
    left, right = report([a]), report([b])
    for payload in (left, right):
        payload["summary"] = {"count": 1, "slow_threshold_seconds": 1.0}
    assert _visibility("slow_report.json", left, right)["items"] == [b, a]
    right["summary"]["slow_threshold_seconds"] = 2.0
    with pytest.raises(ValueError, match="different slow-test thresholds"):
        _visibility("slow_report.json", left, right)


@pytest.mark.parametrize("name", ["not_collected_report.json", "missing_marker_report.json"])
def test_collection_wide_diagnostics_must_match_and_count_once(name: str) -> None:
    left = {
        "generated_at": "before",
        "items": [],
        "summary": {"count": 0},
        "collection_dependencies": {"python-docx": True},
        "pytest": {"selection": "full", "xdist": {"numprocesses": 4, "maxprocesses": 6}},
    }
    right = {
        **left,
        "generated_at": "later",
        "pytest": {"selection": "full", "xdist": {"numprocesses": 2, "maxprocesses": 6}},
    }
    result = _visibility(name, left, right)
    assert result["summary"] == {"count": 0}
    assert result["pytest"]["xdist"] == {"numprocesses": 6, "maxprocesses": 12}
    with pytest.raises(ValueError, match="disagree"):
        _visibility(name, left, {**right, "items": [{"nodeid": "one"}]})
    gap = {**left, "summary": {"count": 1}, "items": [{"nodeid": "one"}]}
    with pytest.raises(ValueError, match=r"markers missing|requires all"):
        _visibility(name, gap, gap)
