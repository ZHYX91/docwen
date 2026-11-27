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


@pytest.mark.parametrize("group", ["control", "apps-plugins", "remaining"])
def test_shard_trial_exports_the_named_database_and_inventory(tmp_path: Path, group: str) -> None:
    runtime = tmp_path / "runtime"
    reports = runtime / "reports"
    reports.mkdir(parents=True)
    for name in (*qa_reports.VISIBILITY_REPORTS, "node-evidence.json"):
        (reports / name).write_text("{}", encoding="utf-8")
    (runtime / ".coverage").write_bytes(b"database")
    (runtime / "private-input.docx").write_bytes(b"fixture")
    if group == "control":
        (runtime / "coverage.xml").write_bytes(b"<coverage/>")

    output = tmp_path / "export"
    qa_reports.export_reports(runtime, output, coverage=True, exit_code=0, shard=group)
    summary = json.loads((output / "summary.json").read_text(encoding="utf-8"))
    assert summary["shard"] == group
    assert (output / "coverage-data.sqlite").read_bytes() == b"database"
    assert (output / "node-evidence.json").read_text(encoding="utf-8") == "{}"
    assert (output / "coverage.xml").exists() == (group == "control")
    assert not (output / "private-input.docx").exists()
