"""services 单元测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen.services.strategies.layout import operations as layout_ops
from docwen.services.strategies.layout.operations import MergePdfsStrategy, SplitPdfStrategy
from docwen.services.strategies.layout.utils import LayoutPreprocessResult
from docwen.utils import workspace_manager
from docwen.utils.workspace_manager import IntermediateItem

pytestmark = pytest.mark.unit


def _create_pdf(path: Path, pages: int) -> None:
    fitz = pytest.importorskip("fitz")
    doc = fitz.open()
    for _ in range(pages):
        doc.new_page()
    doc.save(str(path))
    doc.close()


def test_merge_pdfs_strategy_merges_pages(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    fitz = pytest.importorskip("fitz")

    pdf1 = tmp_path / "a.pdf"
    pdf2 = tmp_path / "b.pdf"
    _create_pdf(pdf1, pages=1)
    _create_pdf(pdf2, pages=2)

    monkeypatch.setattr(workspace_manager, "get_output_directory", lambda _p: str(tmp_path), raising=True)

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
    out_path = Path(result.output_path)
    assert out_path.exists()

    with fitz.open(str(out_path)) as merged:
        assert merged.page_count == 3


def test_split_pdf_strategy_splits_into_two_files(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    fitz = pytest.importorskip("fitz")

    src = tmp_path / "src.pdf"
    _create_pdf(src, pages=3)

    monkeypatch.setattr(workspace_manager, "get_output_directory", lambda _p: str(tmp_path), raising=True)

    result = SplitPdfStrategy().execute(
        str(src),
        options={"pages": [1], "actual_format": "pdf"},
    )

    assert result.success is True
    assert result.output_path is not None
    out1 = Path(result.output_path)
    assert out1.exists()

    generated = [p for p in tmp_path.glob("*.pdf") if p.name != src.name]
    assert len(generated) == 2

    page_counts = []
    for p in generated:
        with fitz.open(str(p)) as doc:
            page_counts.append(doc.page_count)

    assert sorted(page_counts) == [1, 2]


def test_split_pdf_strategy_every_page_generates_n_files(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    fitz = pytest.importorskip("fitz")

    src = tmp_path / "src.pdf"
    _create_pdf(src, pages=3)

    monkeypatch.setattr(workspace_manager, "get_output_directory", lambda _p: str(tmp_path), raising=True)

    result = SplitPdfStrategy().execute(
        str(src),
        options={"split_mode": "every_page", "actual_format": "pdf"},
    )

    assert result.success is True
    assert result.output_path is not None
    out1 = Path(result.output_path)
    assert out1.exists()

    generated = [p for p in tmp_path.glob("*.pdf") if p.name != src.name]
    assert len(generated) == 3
    for p in generated:
        with fitz.open(str(p)) as doc:
            assert doc.page_count == 1


def test_split_pdf_strategy_odd_even_generates_two_files(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    fitz = pytest.importorskip("fitz")

    src = tmp_path / "src.pdf"
    _create_pdf(src, pages=3)

    monkeypatch.setattr(workspace_manager, "get_output_directory", lambda _p: str(tmp_path), raising=True)

    result = SplitPdfStrategy().execute(
        str(src),
        options={"split_mode": "odd_even", "actual_format": "pdf"},
    )

    assert result.success is True
    assert result.output_path is not None
    out1 = Path(result.output_path)
    assert out1.exists()

    generated = [p for p in tmp_path.glob("*.pdf") if p.name != src.name]
    assert len(generated) == 2

    page_counts = []
    for p in generated:
        with fitz.open(str(p)) as doc:
            page_counts.append(doc.page_count)

    assert sorted(page_counts) == [1, 2]


def test_split_pdf_strategy_every_page_single_page_returns_failure(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    pytest.importorskip("fitz")

    src = tmp_path / "src.pdf"
    _create_pdf(src, pages=1)

    monkeypatch.setattr(workspace_manager, "get_output_directory", lambda _p: str(tmp_path), raising=True)

    result = SplitPdfStrategy().execute(
        str(src),
        options={"split_mode": "every_page", "actual_format": "pdf"},
    )

    assert result.success is False


def test_split_pdf_strategy_odd_even_single_page_returns_failure(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    pytest.importorskip("fitz")

    src = tmp_path / "src.pdf"
    _create_pdf(src, pages=1)

    monkeypatch.setattr(workspace_manager, "get_output_directory", lambda _p: str(tmp_path), raising=True)

    result = SplitPdfStrategy().execute(
        str(src),
        options={"split_mode": "odd_even", "actual_format": "pdf"},
    )

    assert result.success is False


# ============================================================
# OFD / CAJ 输入的边界行为
# ============================================================


def test_merge_pdfs_strategy_rejects_caj_early(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """CAJ 输入应在合并入口被提前拦截，不走 preprocess 流程"""
    pytest.importorskip("fitz")

    pdf1 = tmp_path / "a.pdf"
    caj = tmp_path / "b.caj"
    _create_pdf(pdf1, pages=1)
    caj.write_bytes(b"fake-caj")

    monkeypatch.setattr(workspace_manager, "get_output_directory", lambda _p: str(tmp_path), raising=True)

    result = MergePdfsStrategy().execute(
        str(pdf1),
        options={
            "file_list": [str(pdf1), str(caj)],
            "actual_formats": ["pdf", "caj"],
            "selected_file": str(pdf1),
        },
    )

    assert result.success is False
    assert "CAJ" in result.message


def test_split_pdf_strategy_rejects_caj_early(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """CAJ 输入应在拆分入口被提前拦截"""
    caj = tmp_path / "src.caj"
    caj.write_bytes(b"fake-caj")

    monkeypatch.setattr(workspace_manager, "get_output_directory", lambda _p: str(tmp_path), raising=True)

    result = SplitPdfStrategy().execute(
        str(caj),
        options={"pages": [1], "actual_format": "caj"},
    )

    assert result.success is False
    assert "CAJ" in result.message


def test_merge_pdfs_strategy_with_ofd_input_preprocess_success(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """OFD 输入通过 preprocess 转为 PDF 后正常合并"""
    fitz = pytest.importorskip("fitz")

    pdf1 = tmp_path / "a.pdf"
    ofd = tmp_path / "b.ofd"
    _create_pdf(pdf1, pages=1)
    ofd.write_bytes(b"fake-ofd")

    # 准备一个 2 页的 PDF 作为 OFD 预处理的输出
    fake_converted = tmp_path / "converted_ofd.pdf"
    _create_pdf(fake_converted, pages=2)

    def fake_preprocess(file_path, temp_dir, cancel_event=None, actual_format=None):
        """mock: OFD → 直接返回预生成的 PDF；PDF → 复制到 temp_dir"""
        import shutil

        if actual_format == "ofd":
            dst = Path(temp_dir) / "input.pdf"
            shutil.copy2(str(fake_converted), str(dst))
            return LayoutPreprocessResult(
                processed_file=str(dst),
                actual_format="ofd",
                intermediates=[IntermediateItem(kind="layout_pdf", path=str(dst))],
                preprocess_chain=["ofd", "pdf"],
            )
        # PDF 走正常复制
        dst = Path(temp_dir) / "input.pdf"
        shutil.copy2(file_path, str(dst))
        return LayoutPreprocessResult(
            processed_file=str(dst),
            actual_format="pdf",
            preprocess_chain=["pdf"],
        )

    monkeypatch.setattr(layout_ops, "preprocess_layout_file", fake_preprocess, raising=True)
    monkeypatch.setattr(workspace_manager, "get_output_directory", lambda _p: str(tmp_path), raising=True)

    result = MergePdfsStrategy().execute(
        str(pdf1),
        options={
            "file_list": [str(pdf1), str(ofd)],
            "actual_formats": ["pdf", "ofd"],
            "selected_file": str(pdf1),
        },
    )

    assert result.success is True
    assert result.output_path is not None
    with fitz.open(result.output_path) as merged:
        assert merged.page_count == 3  # 1 + 2


def test_merge_pdfs_strategy_with_ofd_preprocess_failure(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """OFD 预处理失败（如 easyofd 不可用）时，合并应返回一致的错误消息"""
    pytest.importorskip("fitz")

    pdf1 = tmp_path / "a.pdf"
    ofd = tmp_path / "b.ofd"
    _create_pdf(pdf1, pages=1)
    ofd.write_bytes(b"fake-ofd")

    call_count = 0

    def fake_preprocess(file_path, temp_dir, cancel_event=None, actual_format=None):
        nonlocal call_count
        import shutil

        call_count += 1
        if actual_format == "ofd":
            raise RuntimeError("OFD转PDF需要easyofd库。\n请安装: pip install easyofd")
        dst = Path(temp_dir) / "input.pdf"
        shutil.copy2(file_path, str(dst))
        return LayoutPreprocessResult(
            processed_file=str(dst),
            actual_format="pdf",
            preprocess_chain=["pdf"],
        )

    monkeypatch.setattr(layout_ops, "preprocess_layout_file", fake_preprocess, raising=True)
    monkeypatch.setattr(workspace_manager, "get_output_directory", lambda _p: str(tmp_path), raising=True)

    result = MergePdfsStrategy().execute(
        str(pdf1),
        options={
            "file_list": [str(pdf1), str(ofd)],
            "actual_formats": ["pdf", "ofd"],
            "selected_file": str(pdf1),
        },
    )

    assert result.success is False
    assert "easyofd" in result.message


def test_split_pdf_strategy_with_ofd_preprocess_failure(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """OFD 拆分时预处理失败应返回一致的错误消息"""
    ofd = tmp_path / "src.ofd"
    ofd.write_bytes(b"fake-ofd")

    def fake_preprocess(file_path, temp_dir, cancel_event=None, actual_format=None):
        raise RuntimeError("OFD转PDF需要easyofd库。\n请安装: pip install easyofd")

    monkeypatch.setattr(layout_ops, "preprocess_layout_file", fake_preprocess, raising=True)
    monkeypatch.setattr(workspace_manager, "get_output_directory", lambda _p: str(tmp_path), raising=True)

    result = SplitPdfStrategy().execute(
        str(ofd),
        options={"pages": [1], "actual_format": "ofd"},
    )

    assert result.success is False
    assert "easyofd" in result.message
