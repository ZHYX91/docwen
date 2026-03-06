"""converter 单元测试。"""

from __future__ import annotations

import os
import shutil
import threading
from pathlib import Path

import pytest

import docwen.converter.formats.layout.external as external
from docwen.converter.formats.layout.external import pdf_to_docx

pytestmark = pytest.mark.unit


def test_pdf_to_docx_falls_back_when_target_move_fails(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    base_dir = tmp_path / "base_dir"
    base_dir.mkdir()
    bad_out = tmp_path / "bad_out"
    bad_out.mkdir()

    src = base_dir / "a.pdf"
    src.write_bytes(b"%PDF-1.4\n%stub\n")

    target = bad_out / "out.docx"

    def _stub_convert_pdf_with_word(pdf_path: str, output_path: str, cancel_event=None):
        Path(output_path).write_bytes(b"docx")
        return output_path

    monkeypatch.setattr(external, "_convert_pdf_with_word", _stub_convert_pdf_with_word, raising=True)
    monkeypatch.setattr(external, "_convert_pdf_with_libreoffice", lambda *_a, **_k: None, raising=True)
    monkeypatch.setattr(external, "_convert_pdf_with_pdf2docx", lambda *_a, **_k: None, raising=True)

    def _stub_move_file_with_retry(source: str, destination: str, *_args, **_kwargs):
        if os.path.normcase(destination).startswith(os.path.normcase(str(bad_out))):
            return None
        Path(destination).parent.mkdir(parents=True, exist_ok=True)
        shutil.move(source, destination)
        return destination

    monkeypatch.setattr("docwen.utils.workspace_manager.move_file_with_retry", _stub_move_file_with_retry, raising=True)

    out = pdf_to_docx(str(src), str(target), cancel_event=threading.Event(), headless=True)
    assert out is not None
    out_path = Path(out)
    assert out_path.exists() is True
    assert os.path.normcase(str(base_dir)) in os.path.normcase(str(out_path.parent))
