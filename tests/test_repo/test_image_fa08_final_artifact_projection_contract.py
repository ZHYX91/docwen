"""Fail-closed guards for the finite FA-08 image final-artifact evidence."""

from __future__ import annotations

from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
FIXTURE_PATH = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_image_format_semantics.json"
REPORT_NAME = "image-final-artifact-parity-2026-07-22.md"


def test_fa08_current_only_sidecar_fix_and_probe_are_guarded() -> None:
    converter = (
        PROJECT_ROOT / "packages" / "plugins" / "image" / "src" / "docwen_plugin_image" / "to_markdown" / "converter.py"
    ).read_text(encoding="utf-8")
    tests = read_source_text(PROJECT_ROOT / "packages" / "plugins" / "image" / "tests" / "test_markdown_assets_*.py")
    probe = (PROJECT_ROOT / "tools" / "validation" / "probe_image_fa08_parity.py").read_text(encoding="utf-8")

    assert 'if enable_ocr and ocr_placement_mode == "image_md":' in converter
    assert "test_image_md_placement_empty_ocr_still_creates_auxiliary" in tests
    assert 'assert [artifact.kind for artifact in result.artifacts] == ["primary", "image", "auxiliary"]' in tests
    for token in (
        'PROJECTS = ("docwen-ref-tk", "docwen-ref-pyside6", "docwen-current")',
        'ROUTES = ("jpg", "webp", "pdf-original", "markdown-file", "markdown-image-md")',
        '"minimum_display_oriented_rgb_similarity"',
        'parser.add_argument("--reproject-only", action="store_true")',
        'parser.add_argument("--rerun-route", action="append", default=[])',
        'result["pass"] = all(result["acceptance"].values())',
    ):
        assert token in probe
