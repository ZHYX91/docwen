"""converter 单元测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen.converter.md2docx import core as md2docx_core

pytestmark = pytest.mark.unit


def test_process_docx_template_removes_temp_file_when_move_fails(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    temp_docx = tmp_path / "temp.docx"
    temp_docx.write_text("x", encoding="utf-8")

    def fake_replace_placeholders(*_args, **_kwargs):
        return str(temp_docx), []

    monkeypatch.setattr(md2docx_core.docx_processor, "replace_placeholders", fake_replace_placeholders, raising=True)
    monkeypatch.setattr(md2docx_core, "move_file_with_retry", lambda *_args, **_kwargs: None, raising=True)

    ok = md2docx_core._process_docx_template(
        doc=None,
        output_path=str(tmp_path / "out.docx"),
        yaml_data={},
        body_data=[],
        template_name=None,
        footnotes=None,
        endnotes=None,
    )

    assert ok is False
    assert temp_docx.exists() is False
