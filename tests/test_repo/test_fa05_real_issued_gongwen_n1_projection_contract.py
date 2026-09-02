"""Fail-closed evidence guards for VIS-174/VIS-191 finite-contract FA-05."""

from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

PROJECT_ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "fa05-real-issued-gongwen-n1-stage-card-2026-07-22.md"
CLOSURE_REPORT_NAME = "fa05-visible-schema-and-reference-defect-closure-2026-07-23.md"


def _read(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def test_fa05_current_image_and_attachment_repairs_have_direct_executable_guards() -> None:
    manifest = _read(
        PROJECT_ROOT
        / "packages"
        / "plugins"
        / "optimizers"
        / "gongwen"
        / "src"
        / "docwen_plugin_optimizer_gongwen"
        / "manifest.py"
    )
    pipeline = _read(
        PROJECT_ROOT
        / "packages"
        / "plugins"
        / "optimizers"
        / "gongwen"
        / "src"
        / "docwen_plugin_optimizer_gongwen"
        / "pipeline.py"
    )
    plugin = _read(
        PROJECT_ROOT
        / "packages"
        / "plugins"
        / "optimizers"
        / "gongwen"
        / "src"
        / "docwen_plugin_optimizer_gongwen"
        / "plugin.py"
    )
    renderer = _read(
        PROJECT_ROOT
        / "packages"
        / "plugins"
        / "optimizers"
        / "gongwen"
        / "src"
        / "docwen_plugin_optimizer_gongwen"
        / "rendering"
        / "markdown_renderer.py"
    )
    rendering_tests = _read(
        PROJECT_ROOT / "packages" / "plugins" / "optimizers" / "gongwen" / "tests" / "test_gongwen_rendering.py"
    )
    golden_tests = _read(
        PROJECT_ROOT / "packages" / "plugins" / "optimizers" / "gongwen" / "tests" / "test_gongwen_golden.py"
    )

    for option in ("to_md_keep_images", "image_mode", "image_link_style"):
        assert option in manifest
        assert option in pipeline
    assert 'options["output_dir"] = str(context.workspace.staging_dir)' in plugin
    assert "ARTIFACT_KIND_IMAGE" in plugin
    assert 'gongwen_result.get("image_paths", [])' in plugin
    assert "generate_image_markdown" in renderer
    assert "Path(img_path).name" in renderer
    assert "_attachment_line_text" in pipeline
    assert "test_skipped_structural_paragraph_preserves_relative_image_reference" in rendering_tests
    assert "test_attachment_line_preserves_original_prefix_when_numbering_is_kept" in rendering_tests
    assert "test_gongwen_image_artifact_finalizes_with_relative_markdown_link" in golden_tests
