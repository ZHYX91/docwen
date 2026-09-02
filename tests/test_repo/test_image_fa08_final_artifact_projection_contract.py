"""Current source and golden guards for the FA-08 image routes."""

from __future__ import annotations

import json
from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
FIXTURE_PATH = PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_image_format_semantics.json"


def test_fa08_current_only_sidecar_fix_is_guarded_by_behavior_and_fixture() -> None:
    converter = (
        PROJECT_ROOT / "packages" / "plugins" / "image" / "src" / "docwen_plugin_image" / "to_markdown" / "converter.py"
    ).read_text(encoding="utf-8")
    tests = read_source_text(PROJECT_ROOT / "packages" / "plugins" / "image" / "tests" / "test_markdown_assets_*.py")
    fixture = json.loads(FIXTURE_PATH.read_text(encoding="utf-8"))
    repair = fixture["fa08_final_artifact_contract_addendum"]["current_only_repair"]

    assert 'if enable_ocr and ocr_placement_mode == "image_md":' in converter
    assert "test_image_md_placement_empty_ocr_still_creates_auxiliary" in tests
    assert 'assert [artifact.kind for artifact in result.artifacts] == ["primary", "image", "auxiliary"]' in tests
    assert repair["classification"] == "fixed_empty_ocr_image_md_companion_parity"
    assert repair["affected_slots"] == ["B1/markdown-image-md", "B2/markdown-image-md"]
    assert repair["regression_test"].endswith(
        "TestImageToMarkdownOcrPlacementImageMd::test_image_md_placement_empty_ocr_still_creates_auxiliary"
    )
