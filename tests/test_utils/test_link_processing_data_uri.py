"""utils 单元测试。"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from docwen.utils.link_processing import process_markdown_links

pytestmark = pytest.mark.unit


_BASE64_PNG_1X1 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg=="


def _extract_image_path(result: str) -> str:
    m = re.search(r"\{\{IMAGE:([^}]+)\}\}", result)
    assert m, result
    payload = m.group(1)
    return payload.split(r"\|", 1)[0]


def test_markdown_data_uri_image_embeds_as_placeholder_and_writes_temp_file(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setattr("docwen.utils.link_processing.config_manager.get_max_embed_depth", lambda: 10)
    monkeypatch.setattr("docwen.utils.link_processing.config_manager.get_markdown_embed_image_mode", lambda: "embed")
    monkeypatch.setattr("docwen.utils.link_processing.config_manager.get_wiki_embed_image_mode", lambda: "embed")
    monkeypatch.setattr("docwen.utils.link_processing.config_manager.get_wiki_link_mode", lambda: "keep")
    monkeypatch.setattr("docwen.utils.link_processing.config_manager.get_markdown_link_mode", lambda: "keep")

    md = f"前文\\n\\n![测试图片](data:image/png;base64,{_BASE64_PNG_1X1})\\n\\n后文\\n"
    source = tmp_path / "src.md"
    source.write_text(md, encoding="utf-8")

    result = process_markdown_links(md, str(source), temp_dir=str(tmp_path))
    assert "{{IMAGE:" in result

    image_path = _extract_image_path(result)
    assert Path(image_path).exists()
    assert Path(image_path).parent == tmp_path
    assert Path(image_path).suffix.lower() == ".png"
    assert Path(image_path).stat().st_size > 0


def test_markdown_data_uri_image_respects_size_syntax(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setattr("docwen.utils.link_processing.config_manager.get_max_embed_depth", lambda: 10)
    monkeypatch.setattr("docwen.utils.link_processing.config_manager.get_markdown_embed_image_mode", lambda: "embed")
    monkeypatch.setattr("docwen.utils.link_processing.config_manager.get_wiki_link_mode", lambda: "keep")
    monkeypatch.setattr("docwen.utils.link_processing.config_manager.get_markdown_link_mode", lambda: "keep")

    md = f"![alt](data:image/png;base64,{_BASE64_PNG_1X1} =300x200)"
    source = tmp_path / "src.md"
    source.write_text(md, encoding="utf-8")

    result = process_markdown_links(md, str(source), temp_dir=str(tmp_path))
    assert r"\|300\|200" in result


def test_markdown_data_uri_invalid_base64_extracts_alt_text(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setattr("docwen.utils.link_processing.config_manager.get_max_embed_depth", lambda: 10)
    monkeypatch.setattr(
        "docwen.utils.link_processing.config_manager.get_markdown_embed_image_mode", lambda: "extract_text"
    )
    monkeypatch.setattr("docwen.utils.link_processing.config_manager.get_wiki_link_mode", lambda: "keep")
    monkeypatch.setattr("docwen.utils.link_processing.config_manager.get_markdown_link_mode", lambda: "keep")

    md = "![替代文本](data:image/png;base64,NOT_BASE64!!!)"
    source = tmp_path / "src.md"
    source.write_text(md, encoding="utf-8")

    result = process_markdown_links(md, str(source), temp_dir=str(tmp_path))
    assert result == "替代文本"


def test_markdown_data_uri_too_large_emits_warning_and_degrades(tmp_path: Path, monkeypatch, caplog) -> None:
    monkeypatch.setattr("docwen.utils.link_processing.config_manager.get_max_embed_depth", lambda: 10)
    monkeypatch.setattr(
        "docwen.utils.link_processing.config_manager.get_markdown_embed_image_mode", lambda: "extract_text"
    )
    monkeypatch.setattr("docwen.utils.link_processing.config_manager.get_wiki_link_mode", lambda: "keep")
    monkeypatch.setattr("docwen.utils.link_processing.config_manager.get_markdown_link_mode", lambda: "keep")

    large_payload = "A" * (14 * 1024 * 1024)
    md = f"![替代文本](data:image/png;base64,{large_payload})"
    source = tmp_path / "src.md"
    source.write_text(md, encoding="utf-8")

    with caplog.at_level("WARNING"):
        result = process_markdown_links(md, str(source), temp_dir=str(tmp_path))
    assert "超过大小上限" in caplog.text
    assert result == "替代文本"


def test_markdown_data_uri_keep_mode_keeps_original_link(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setattr("docwen.utils.link_processing.config_manager.get_max_embed_depth", lambda: 10)
    monkeypatch.setattr("docwen.utils.link_processing.config_manager.get_markdown_embed_image_mode", lambda: "keep")
    monkeypatch.setattr("docwen.utils.link_processing.config_manager.get_wiki_link_mode", lambda: "keep")
    monkeypatch.setattr("docwen.utils.link_processing.config_manager.get_markdown_link_mode", lambda: "keep")

    md = "![替代文本](data:image/png;base64,NOT_BASE64!!!)"
    source = tmp_path / "src.md"
    source.write_text(md, encoding="utf-8")

    result = process_markdown_links(md, str(source), temp_dir=str(tmp_path))
    assert result == md


def test_markdown_data_uri_remove_mode_removes_link(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setattr("docwen.utils.link_processing.config_manager.get_max_embed_depth", lambda: 10)
    monkeypatch.setattr("docwen.utils.link_processing.config_manager.get_markdown_embed_image_mode", lambda: "remove")
    monkeypatch.setattr("docwen.utils.link_processing.config_manager.get_wiki_link_mode", lambda: "keep")
    monkeypatch.setattr("docwen.utils.link_processing.config_manager.get_markdown_link_mode", lambda: "keep")

    md = "![替代文本](data:image/png;base64,NOT_BASE64!!!)"
    source = tmp_path / "src.md"
    source.write_text(md, encoding="utf-8")

    result = process_markdown_links(md, str(source), temp_dir=str(tmp_path))
    assert result == ""


def test_wiki_data_uri_embed_uses_wiki_mode_and_writes_temp_file(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setattr("docwen.utils.link_processing.config_manager.get_max_embed_depth", lambda: 10)
    monkeypatch.setattr("docwen.utils.link_processing.config_manager.get_wiki_embed_image_mode", lambda: "embed")
    monkeypatch.setattr("docwen.utils.link_processing.config_manager.get_wiki_link_mode", lambda: "keep")
    monkeypatch.setattr("docwen.utils.link_processing.config_manager.get_markdown_link_mode", lambda: "keep")

    md = f"![[data:image/png;base64,{_BASE64_PNG_1X1}]]"
    source = tmp_path / "src.md"
    source.write_text(md, encoding="utf-8")

    result = process_markdown_links(md, str(source), temp_dir=str(tmp_path))
    assert "{{IMAGE:" in result
    image_path = _extract_image_path(result)
    assert Path(image_path).exists()
    assert Path(image_path).parent == tmp_path
