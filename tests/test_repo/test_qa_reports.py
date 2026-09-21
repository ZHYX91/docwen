from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pytest
from tools import qa_reports

pytestmark = pytest.mark.unit


def test_compact_coverage_retains_totals_and_source_identity(tmp_path: Path) -> None:
    runtime = tmp_path / "raw"
    reports = runtime / "reports"
    reports.mkdir(parents=True)
    for name in qa_reports.VISIBILITY_REPORTS:
        (reports / name).write_text("{}")
    xml = b'<coverage lines-valid="100" lines-covered="90"><packages><package name="core" line-rate="0.9"/></packages></coverage>'
    (runtime / "coverage.xml").write_bytes(xml)
    output = tmp_path / "compact"
    qa_reports.export_reports(runtime, output, coverage=True, exit_code=0, coverage_detail=False)
    assert not (output / "coverage.xml").exists()
    summary = json.loads((output / "coverage-summary.json").read_text())
    assert summary["sourceSha256"] == hashlib.sha256(xml).hexdigest()
    assert summary["totals"]["lines-covered"] == "90"
    assert summary["packages"][0]["name"] == "core"
    assert "coverage-summary.json" in json.loads((output / "summary.json").read_text())["reports"]
    assert (runtime / "coverage.xml").read_bytes() == xml
