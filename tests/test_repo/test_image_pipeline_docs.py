"""Current image-export documentation and manifest guards."""

from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

ROOT = Path(__file__).resolve().parents[2]


def test_markdown_compatibility_documents_declared_image_modes() -> None:
    from docwen_plugin_layout.manifest import LAYOUT_TO_MD_OPTIONS_SCHEMA

    text = (ROOT / "docs" / "specs" / "markdown-compatibility.md").read_text(encoding="utf-8")
    image_mode = LAYOUT_TO_MD_OPTIONS_SCHEMA["properties"]["image_mode"]
    assert image_mode["enum"] == ["file", "base64", "embed", "omit"]
    folded = text.casefold()
    for token in ("layout", "presentation", "markup", "file", "base64", "embed", "omit"):
        assert token in folded


def test_layout_image_export_regressions_remain_executable() -> None:
    converter = (ROOT / "packages/plugins/layout/src/docwen_plugin_layout/to_markdown/converter.py").read_text(
        encoding="utf-8"
    )
    tests = "\n".join(
        path.read_text(encoding="utf-8")
        for path in sorted((ROOT / "packages/plugins/layout/tests").glob("test_layout_conversions_*.py"))
    )

    assert "get_markdown_export_modes(" in converter
    assert "generate_image_markdown(" in converter
    assert "test_pdf_to_md_base64_image_mode_inlines_images" in tests
    assert "test_pdf_to_md_omit_image_mode_removes_image_artifacts" in tests
