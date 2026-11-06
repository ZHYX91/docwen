"""Missing Markdown embeds are retained in Word without embedding English errors."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter

from ._link_processing_routes_support import _link_config
from .conftest import make_context

pytestmark = pytest.mark.contract


def test_missing_embeds_issue_one_warning_without_file_not_found_body(tmp_path: Path) -> None:
    source = tmp_path / "source.md"
    source.write_text(
        "# A document\n\nMissing ![[not-here.md]] and ![image](not-here.png)\n",
        encoding="utf-8",
    )
    context, _workspace = make_context(
        str(source), target_format="docx", config_values=_link_config(file_not_found_mode="keep")
    )
    result = MdToDocxConverter().convert(context)
    assert result.success, result.error
    warning = [d for d in result.diagnostics if d.code == "MD2DOCX-EMBED-SOURCE-MISSING"]
    assert len(warning) == 1
    assert warning[0].level == "warning"
    output = Path(result.artifacts[0].staging_path)
    with zipfile.ZipFile(output) as docx:
        body = docx.read("word/document.xml").decode("utf-8")
    assert "File not found" not in body
    assert "not-here.md" in body and "not-here.png" in body


def test_literal_image_syntax_does_not_report_missing_embed(tmp_path: Path) -> None:
    source = tmp_path / "source.md"
    source.write_text("`![literal](missing.png)`\n\n```md\n![[missing.md]]\n```\n", encoding="utf-8")
    context, _workspace = make_context(
        str(source), target_format="docx", config_values=_link_config(file_not_found_mode="keep")
    )
    result = MdToDocxConverter().convert(context)
    assert result.success, result.error
    assert all(d.code != "MD2DOCX-EMBED-SOURCE-MISSING" for d in result.diagnostics)
