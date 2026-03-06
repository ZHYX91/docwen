"""docx_spell 单元测试。"""

from __future__ import annotations

import os
import shutil
from pathlib import Path

import pytest
from docx import Document

from docwen.docx_spell.core import process_docx

pytestmark = pytest.mark.unit


def test_process_docx_uniquifies_on_timestamp_collision(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr("docwen.utils.path_utils.generate_timestamp", lambda: "20200101_000000", raising=True)

    src = tmp_path / "input.docx"
    Document().save(src)

    out1 = process_docx(str(src), output_dir=str(tmp_path), proofread_options={})
    out2 = process_docx(str(src), output_dir=str(tmp_path), proofread_options={})

    assert out1 is not None
    assert out2 is not None
    assert Path(out1).exists() is True
    assert Path(out2).exists() is True
    assert out1 != out2
    assert Path(out2).stem.endswith("_001") is True


def test_process_docx_falls_back_when_target_move_fails(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr("docwen.utils.path_utils.generate_timestamp", lambda: "20200101_000000", raising=True)

    base_dir = tmp_path / "base_dir"
    base_dir.mkdir()
    bad_out = tmp_path / "bad_out"
    bad_out.mkdir()

    src = base_dir / "input.docx"
    Document().save(src)

    def _stub_move_file_with_retry(source: str, destination: str, *_args, **_kwargs):
        if os.path.normcase(destination).startswith(os.path.normcase(str(bad_out))):
            return None
        Path(destination).parent.mkdir(parents=True, exist_ok=True)
        shutil.move(source, destination)
        return destination

    monkeypatch.setattr("docwen.utils.workspace_manager.move_file_with_retry", _stub_move_file_with_retry, raising=True)

    out = process_docx(str(src), output_dir=str(bad_out), proofread_options={})
    assert out is not None
    assert Path(out).exists() is True
    assert os.path.normcase(str(base_dir)) in os.path.normcase(str(Path(out).parent))
