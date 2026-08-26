"""Fail-closed contracts for VIS-2026-07-17-127 rich Presentation evidence."""

from __future__ import annotations

import json
from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]
GOLDEN = ROOT / "tests" / "fixtures" / "golden"
FIXTURE = GOLDEN / "old_system_pptx_to_markdown_semantics.json"
REPORT_NAME = "apache-poi-rich-presentation-physical-matrix-2026-07-17.md"
CONVERTER = ROOT / "packages/plugins/presentation/src/docwen_plugin_presentation/pptx_md/converter.py"
PRESENTATION_TEST = ROOT / "packages/plugins/presentation/tests/test_presentation_to_md_*.py"


def _addendum() -> dict[str, object]:
    fixture = json.loads(FIXTURE.read_text(encoding="utf-8"))
    return fixture["real_world_rich_presentation_addendum"]


def test_official_source_identity_and_rich_package_projection_are_pinned() -> None:
    addendum = _addendum()
    assert addendum["evidence_id"] == "VIS-2026-07-17-127"
    source = addendum["source"]
    assert source["name"] == "60810.pptx"
    assert source["repository"] == "apache/poi"
    assert source["commit"] == "86e967d9b28d6a322a87ae8fcbf2a7eeb56cef96"
    assert source["git_blob"] == "3e4e22be5f4911be6d9935246d32f5f76907ba4f"
    assert (source["bytes"], source["sha256"]) == (
        874_522,
        "61afceb0365523ceba6fc00a525157148c2ddbdb3b5d843992ac3c107e9921bf",
    )
    assert source["title"] == "3.7.1 HRBP Process overview"
    assert source["slide_count"] == 28
    assert source["hidden_slide_indexes"] == [4, 21, 22]
    assert source["section_names"] == ["Core layouts"]
    assert (
        source["notes_part_count"],
        source["nonempty_notes_count"],
        source["nonempty_notes_chars"],
    ) == (17, 6, 5_544)
    assert (
        source["package_media_count"],
        source["picture_occurrence_count"],
        source["diagram_part_count"],
    ) == (17, 19, 20)
    assert source["python_pptx_accessible_text_count"] == 60
    assert source["smartart_text_count"] == 11


def test_three_project_final_markdown_and_resource_projection_is_exact() -> None:
    addendum = _addendum()
    execution = addendum["execution"]
    assert execution["route"] == "pptx->md"
    assert execution["success_count"] == execution["expected_success_count"] == 3
    assert execution["conversion_process_added"] == []
    assert execution["conversion_process_removed"] == []

    outputs = addendum["outputs"]
    assert outputs["docwen-ref-tk"]["markdown_sha256"] == (outputs["docwen-ref-pyside6"]["markdown_sha256"])
    for project in ("docwen-ref-tk", "docwen-ref-pyside6", "docwen-current"):
        projection = outputs[project]
        assert projection["h2_count"] == 28
        assert projection["notes_block_count"] == 6
        assert projection["image_link_count"] == projection["image_artifact_count"] == 19
    assert outputs["docwen-current"]["diagnostic_codes"] == [
        "PPTX2MD-OK",
        "FINALIZER_DONE",
    ]

    comparison = addendum["comparison"]
    assert comparison["old_markdown_raw_exact"] is True
    assert comparison["semantic_line_count"] == 294
    assert comparison["semantic_sha256"] == ("e61f6828a791425b3ba57ed7e650898fb6bd1b43718cf32af22d9667bbbe8d9c")
    assert comparison["three_project_semantic_lines_exact_after_documented_normalization"] is True
    assert comparison["three_project_resource_bytes_exact"] is True
    assert comparison["source_accessible_text_matches"] == {
        "docwen-ref-tk": 60,
        "docwen-ref-pyside6": 60,
        "docwen-current": 60,
    }
    assert comparison["source_notes_matches"] == {
        "docwen-ref-tk": 6,
        "docwen-ref-pyside6": 6,
        "docwen-current": 6,
    }


def test_current_jpeg_mime_fix_uses_canonical_core_fact_source() -> None:
    fix = _addendum()["current_fix"]
    assert fix["classification"] == "real_current_only_artifact_metadata_defect_fixed"
    assert fix["before"].endswith("image/jpg")
    assert fix["after"].endswith("image/jpeg")
    assert _addendum()["outputs"]["docwen-current"]["image_media_types"] == {
        "image/jpeg": 4,
        "image/png": 15,
    }

    converter = CONVERTER.read_text(encoding="utf-8")
    assert "from docwen_core.formats.categories import get_media_type" in converter
    assert "media_type=get_media_type(ext)" in converter
    assert 'media_type=f"image/{ext}"' not in converter
    tests = read_source_text(PRESENTATION_TEST)
    assert "def test_pptx_jpeg_image_uses_standard_media_type" in tests
    assert 'assert image_artifact.media_type == "image/jpeg"' in tests
