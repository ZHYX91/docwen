"""YAML front-matter ordinary-link parity for Markdown -> DOCX."""

from __future__ import annotations

from pathlib import Path

import pytest
from docx import Document

from ._link_processing_routes_support import _convert_docx, _link_config

pytestmark = pytest.mark.contract


def _gongwen_config(*, wiki_mode: str, markdown_mode: str) -> dict:
    config = _link_config(wiki_mode=wiki_mode, markdown_mode=markdown_mode)
    config["gui"] = {"language": {"locale": "zh_CN"}}
    config["field_processors"] = {
        "settings": {"order": ["gongwen"]},
        "processors": {
            "gongwen": {
                "module": "docwen_plugin_markdown.field_processors.gongwen",
                "locales": ["zh_CN"],
                "enabled": True,
            }
        },
    }
    return config


def _template(tmp_path: Path) -> Path:
    path = tmp_path / "yaml-links-template.docx"
    doc = Document()
    doc.paragraphs[0].text = "抄送：{{抄送机关}}"
    doc.add_paragraph("网站：{{site}}")
    doc.add_paragraph("{{正文}}")
    doc.save(path)
    return path


def _source(tmp_path: Path) -> Path:
    (tmp_path / "guide.md").write_text("target", encoding="utf-8")
    source = tmp_path / "yaml-links.md"
    source.write_text(
        "---\n"
        "抄送机关:\n"
        "  - \"[[guide|喵喵]]\"\n"
        "  - 钱钱钱\n"
        "site: \"[项目主页](https://example.com/project)\"\n"
        "份号: \"001\"\n"
        "enabled: false\n"
        "count: 0\n"
        "---\n\n"
        "正文。\n",
        encoding="utf-8",
    )
    return source


@pytest.mark.parametrize("mode", ["keep", "extract_text", "remove", "hyperlink"])
def test_yaml_links_follow_request_scoped_ordinary_link_policy(tmp_path: Path, mode: str) -> None:
    source = _source(tmp_path)
    template = _template(tmp_path)
    observation = _convert_docx(
        source,
        _gongwen_config(wiki_mode=mode, markdown_mode=mode),
        options={"template_name": str(template)},
    )

    if mode == "keep":
        assert "[[guide|喵喵]]" in observation.text
        assert "[项目主页](https://example.com/project)" in observation.text
        assert not observation.hyperlink_targets
    elif mode == "extract_text":
        assert "抄送：喵喵，钱钱钱" in observation.text
        assert "网站：项目主页" in observation.text
        assert "[[guide|喵喵]]" not in observation.text
        assert "[项目主页](" not in observation.text
        assert not observation.hyperlink_targets
    elif mode == "remove":
        assert "抄送：钱钱钱" in observation.text
        assert "喵喵" not in observation.text
        assert "项目主页" not in observation.text
        assert not observation.hyperlink_targets
    else:
        assert "抄送：喵喵，钱钱钱" in observation.text
        assert "网站：项目主页" in observation.text
        assert "https://example.com/project" in observation.hyperlink_targets
        assert any(
            target.startswith("file:") and target.endswith("/guide.md") for target in observation.hyperlink_targets
        )
        assert "<w:hyperlink" in observation.document_xml
