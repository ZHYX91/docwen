"""Fail-closed evidence guards for VIS-176 / finite-contract FA-12."""

from __future__ import annotations

from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = [pytest.mark.contract, pytest.mark.golden]

PROJECT_ROOT = Path(__file__).resolve().parents[2]
FIXTURE = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_mhtml_to_markdown_semantics.json"
REPORT_NAME = "markup-presentation-real-corpus-final-artifact-parity-2026-07-22.md"
STAGE_CARD = "fa12-real-web-presentation-stage-card-2026-07-22.md"
RECONCILIATION_REPORT = "fa12-final-artifact-reconciliation-2026-07-23.md"
RECONCILIATION_CARD = "fa12-final-artifact-reconciliation-stage-card-2026-07-23.md"
HISTORICAL_STATUS = "CURRENT_N1_FIXED_STRICT_REFERENCE_AND_POLICY_BOUNDARY_UNACCEPTED"
STATUS = "FIXED_AND_VERIFIED"


def _read(path: Path) -> str:
    return read_source_text(path)


def test_fa12_mhtml_repair_has_direct_executable_guards() -> None:
    converter = _read(
        PROJECT_ROOT
        / "packages"
        / "plugins"
        / "markup"
        / "src"
        / "docwen_plugin_markup"
        / "web_archive"
        / "converter.py"
    )
    tests = _read(PROJECT_ROOT / "packages" / "plugins" / "markup" / "tests" / "test_input_routes_*.py")

    assert "_decode_mhtml_html_payload" in converter
    assert "_extract_html_image_sources" in converter
    assert "return unescape(match.group(1).strip())" in converter
    assert "test_mhtml_decodes_title_and_finalizes_only_body_images" in tests
    assert "test_mhtml_uses_html_meta_charset_when_mime_part_omits_it" in tests
