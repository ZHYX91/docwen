"""md2docx xml_processor 的单元测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen.converter.md2docx.processors import xml_processor

pytestmark = pytest.mark.unit


def test_process_docx_file_outputs_processed_and_uniquifies(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    input_docx = tmp_path / "report.docx"
    input_docx.write_bytes(b"x")

    existing = tmp_path / "report_processed.docx"
    existing.write_text("old", encoding="utf-8")

    monkeypatch.setattr(xml_processor, "extract_docx", lambda *_args, **_kwargs: None, raising=True)
    monkeypatch.setattr(xml_processor, "process_xml_file", lambda *_args, **_kwargs: None, raising=True)
    monkeypatch.setattr(xml_processor, "has_unreplaced_placeholders", lambda *_args, **_kwargs: False, raising=True)

    def fake_repack(_source_dir: str, output_path: str) -> None:
        Path(output_path).write_text("new", encoding="utf-8")

    monkeypatch.setattr(xml_processor, "repack_docx", fake_repack, raising=True)

    out = xml_processor.process_docx_file(str(input_docx), {"a": "b"})

    assert Path(out).name == "report_processed_001.docx"
    assert existing.read_text(encoding="utf-8") == "old"
    assert Path(out).read_text(encoding="utf-8") == "new"


def test_xml_processor_process_xml_file_minimal(tmp_path: Path) -> None:
    xml_path = tmp_path / "doc.xml"
    xml_path.write_text(
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        "<w:body><w:p><w:r><w:t>hello</w:t></w:r></w:p></w:body>"
        "</w:document>",
        encoding="utf-8",
    )

    xml_processor.process_xml_file(str(xml_path), {})
    assert "hello" in xml_path.read_text(encoding="utf-8")
