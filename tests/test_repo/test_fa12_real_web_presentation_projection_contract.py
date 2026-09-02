"""Fail-closed evidence guards for VIS-176 / finite-contract FA-12."""

from __future__ import annotations

from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = [pytest.mark.contract, pytest.mark.golden]

PROJECT_ROOT = Path(__file__).resolve().parents[2]


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
