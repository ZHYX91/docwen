"""Fail-closed evidence guards for VIS-180 / finite-contract FA-13-N1."""

from __future__ import annotations

from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = [pytest.mark.contract, pytest.mark.golden]

PROJECT_ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "fa13-real-german-localized-docx-n1-current-fix-2026-07-23.md"
STAGE_CARD = "fa13-real-german-localized-docx-n1-stage-card-2026-07-23.md"
STATUS = "N1_FIXED_AND_VERIFIED_N2_PENDING_EXTERNAL_SAMPLE_PERMISSION"
FINAL_STATUS = "FIXED_AND_VERIFIED_WITH_USER_ACCEPTED_DERIVED_IMAGE_ACCURACY_BOUNDARY"
SOURCE_SHA = "44A4CB870A0D3E58EF2B1B719543A5899186F543DC92148287B1C357F719310A"
OUTPUT_SHA = "52190B6A9160F88573F4B5FACCFBC454283026F598C379C275AEAED4258F992D"
SUMMARY_SHA = "C0A8161DB47B7646EA677337FEF60F83E34BB2833528065226A4C0471CAD742B"


def _read(path: Path) -> str:
    return read_source_text(path)


def test_fa13_n1_has_direct_mixed_revision_regression() -> None:
    renderer = _read(
        PROJECT_ROOT
        / "packages"
        / "plugins"
        / "document"
        / "src"
        / "docwen_plugin_document"
        / "shared"
        / "markdown_runs.py"
    )
    converter = _read(
        PROJECT_ROOT
        / "packages"
        / "plugins"
        / "document"
        / "src"
        / "docwen_plugin_document"
        / "to_markdown"
        / "converter.py"
    )
    tests = _read(
        PROJECT_ROOT / "packages" / "plugins" / "document" / "tests" / "test_to_markdown_standard_parity_*.py"
    )

    assert "def _flush_text()" in renderer
    assert 'elif tag == "tab"' in renderer
    assert 'for tag in ("ins", "fldSimple", "hyperlink")' in converter
    assert "test_mixed_direct_and_inserted_revision_text_reaches_paragraph_processing" in tests
    assert 'assert "Rejected heading" not in' in tests
