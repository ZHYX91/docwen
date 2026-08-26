"""Guards for the VIS-099 comprehensive Markdown output parity evidence."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
FIXTURE_NAME = "old_system_markdown_output_batch_semantics.json"
REPORT_NAME = "markdown-output-comprehensive-batch-parity-2026-07-16.md"


def _read(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def test_markdown_output_batch_fixture_records_three_project_semantics() -> None:
    fixture_path = PROJECT_ROOT / "tests" / "fixtures" / "golden" / FIXTURE_NAME
    fixture = json.loads(_read(fixture_path))

    assert fixture["golden_id"] == "GOLDEN-001"
    assert set(fixture["projects"]) == {"old_tk", "old_pyside6", "current"}
    assert fixture["source"]["same_logical_bytes_in_all_projects"] is True
    assert fixture["templates"]["docx"]["same_bytes_in_all_projects"] is True
    assert fixture["templates"]["xlsx"]["same_bytes_in_all_projects"] is True
    assert len(fixture["template_contract"]["docx_templates"]) == 11
    assert fixture["shared_docx_projection"]["title"] == "Test File"
    assert fixture["shared_docx_projection"]["visible_paragraph_count"] == 83
    assert fixture["xlsx_projection"]["same_normalized_cells_in_all_projects"] is True
    assert fixture["xlsx_projection"]["cells"]["B1"] == "Test File"
    assert fixture["csv_projection"]["same_normalized_rows_in_all_projects"] is True
    assert fixture["current_finalizer_projection"]["workspace_or_staging_path_leak"] is False
    assert len(fixture["defects_closed"]) == 3
    assert any("not the acceptance oracle" in item for item in fixture["boundary"])
    assert any("overall parity" in item for item in fixture["boundary"])
