"""converter 单元测试。"""

from __future__ import annotations

import os
import shutil
import threading
from pathlib import Path

import pytest

from docwen.converter.formats.pdf_export.document import docx_to_pdf
from docwen.converter.formats.pdf_export.spreadsheet import xlsx_to_pdf

pytestmark = pytest.mark.unit


def _stub_convert_with_fallback_success(*, output_path: str, **_kwargs):
    Path(output_path).write_bytes(b"%PDF-1.4\n%stub\n")
    return output_path, "stub"


def test_docx_to_pdf_uniquifies_on_timestamp_collision(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr("docwen.utils.path_utils.generate_timestamp", lambda: "20200101_000000", raising=True)
    monkeypatch.setattr(
        "docwen.utils.file_type_utils.detect_actual_file_format", lambda *_a, **_k: "docx", raising=True
    )
    monkeypatch.setattr(
        "docwen.converter.formats.pdf_export.document.convert_with_fallback",
        _stub_convert_with_fallback_success,
        raising=True,
    )

    src = tmp_path / "a.docx"
    src.write_bytes(b"x")

    out1 = docx_to_pdf(str(src))
    out2 = docx_to_pdf(str(src))

    assert out1 is not None and out2 is not None
    assert Path(out1).exists() is True
    assert Path(out2).exists() is True
    assert out1 != out2
    assert Path(out2).stem.endswith("_001") is True


def test_xlsx_to_pdf_falls_back_when_target_move_fails(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr("docwen.utils.path_utils.generate_timestamp", lambda: "20200101_000000", raising=True)
    monkeypatch.setattr(
        "docwen.utils.file_type_utils.detect_actual_file_format", lambda *_a, **_k: "xlsx", raising=True
    )
    monkeypatch.setattr(
        "docwen.converter.formats.pdf_export.spreadsheet.convert_with_fallback",
        _stub_convert_with_fallback_success,
        raising=True,
    )

    base_dir = tmp_path / "base_dir"
    base_dir.mkdir()
    bad_out = tmp_path / "bad_out"
    bad_out.mkdir()

    src = base_dir / "a.xlsx"
    src.write_bytes(b"x")

    target = bad_out / "out.pdf"

    def _stub_move_file_with_retry(source: str, destination: str, *_args, **_kwargs):
        if os.path.normcase(destination).startswith(os.path.normcase(str(bad_out))):
            return None
        Path(destination).parent.mkdir(parents=True, exist_ok=True)
        shutil.move(source, destination)
        return destination

    monkeypatch.setattr("docwen.utils.workspace_manager.move_file_with_retry", _stub_move_file_with_retry, raising=True)

    out = xlsx_to_pdf(str(src), output_path=str(target), cancel_event=threading.Event())
    assert out is not None
    out_path = Path(out)
    assert out_path.exists() is True
    assert os.path.normcase(str(base_dir)) in os.path.normcase(str(out_path.parent))
