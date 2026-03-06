"""converter 单元测试。"""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from lxml import etree

from docwen.converter.md2docx.handlers.note_handler import NoteContext, write_notes_to_docx

pytestmark = pytest.mark.unit


def test_write_notes_to_docx_uses_atomic_replace_and_cleans_temp(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    docx_path = tmp_path / "in.docx"

    with zipfile.ZipFile(docx_path, "w") as zf:
        zf.writestr("word/footnotes.xml", b"<root/>")
        zf.writestr("word/endnotes.xml", b"<root/>")
        zf.writestr("word/document.xml", b"<root/>")

    ctx = NoteContext()
    ctx.footnote_elements.append(etree.Element("note"))

    temp_dir = tmp_path / "tempdir"
    temp_dir.mkdir()

    monkeypatch.setattr("tempfile.mkdtemp", lambda: str(temp_dir), raising=True)

    called = {"count": 0}

    import docwen.utils.workspace_manager as wm

    real_replace = wm.replace_file_atomic

    def wrapped_replace(temp_file: str, target_path: str, *, create_backup: bool = False) -> str:
        called["count"] += 1
        return real_replace(temp_file, target_path, create_backup=create_backup)

    monkeypatch.setattr(wm, "replace_file_atomic", wrapped_replace, raising=True)

    write_notes_to_docx(str(docx_path), ctx)

    assert called["count"] == 1
    assert docx_path.exists() is True
    assert temp_dir.exists() is False
