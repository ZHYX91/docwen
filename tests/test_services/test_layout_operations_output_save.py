"""services 单元测试。"""

from __future__ import annotations

import os
import shutil
from pathlib import Path

import pytest

from docwen.services.strategies.layout.operations import MergePdfsStrategy, SplitPdfStrategy

pytestmark = pytest.mark.unit


def _create_pdf(path: Path, pages: int) -> None:
    fitz = pytest.importorskip("fitz")
    doc = fitz.open()
    for _ in range(pages):
        doc.new_page()
    doc.save(str(path))
    doc.close()


def test_merge_pdfs_strategy_falls_back_when_target_move_fails(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    pytest.importorskip("fitz")

    src_dir = tmp_path / "src"
    src_dir.mkdir()
    bad_out = tmp_path / "bad_out"
    bad_out.mkdir()

    pdf1 = src_dir / "a.pdf"
    pdf2 = src_dir / "b.pdf"
    _create_pdf(pdf1, pages=1)
    _create_pdf(pdf2, pages=2)

    monkeypatch.setattr(
        "docwen.utils.workspace_manager.get_output_directory",
        lambda _p: str(bad_out),
        raising=True,
    )
    monkeypatch.setattr("docwen.utils.path_utils.generate_timestamp", lambda: "20200101_000000", raising=True)

    def _stub_move_file_with_retry(source: str, destination: str, *_args, **_kwargs):
        if os.path.normcase(destination).startswith(os.path.normcase(str(bad_out))):
            return None
        Path(destination).parent.mkdir(parents=True, exist_ok=True)
        shutil.move(source, destination)
        return destination

    monkeypatch.setattr("docwen.utils.workspace_manager.move_file_with_retry", _stub_move_file_with_retry, raising=True)

    result = MergePdfsStrategy().execute(
        str(pdf1),
        options={
            "file_list": [str(pdf1), str(pdf2)],
            "actual_formats": ["pdf", "pdf"],
            "selected_file": str(pdf1),
        },
    )

    assert result.success is True
    assert result.output_path is not None
    out = Path(result.output_path)
    assert out.exists() is True
    assert os.path.normcase(str(src_dir)) in os.path.normcase(str(out.parent))


def test_split_pdf_strategy_uniquifies_when_timestamp_collides(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    pytest.importorskip("fitz")

    src = tmp_path / "src.pdf"
    _create_pdf(src, pages=3)

    monkeypatch.setattr(
        "docwen.utils.workspace_manager.get_output_directory",
        lambda _p: str(tmp_path),
        raising=True,
    )
    monkeypatch.setattr("docwen.utils.path_utils.generate_timestamp", lambda: "20200101_000000", raising=True)

    s = SplitPdfStrategy()
    r1 = s.execute(str(src), options={"pages": [1], "actual_format": "pdf"})
    r2 = s.execute(str(src), options={"pages": [1], "actual_format": "pdf"})

    assert r1.success is True
    assert r2.success is True

    outputs = [p for p in tmp_path.glob("*.pdf") if p.name != src.name]
    assert len(outputs) == 4
    assert any("_001" in p.stem for p in outputs) is True
