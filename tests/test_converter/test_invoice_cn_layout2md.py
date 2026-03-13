"""converter 单元测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen.converter.layout2md.invoice_cn import convert_invoice_cn_layout_to_md
from docwen.translation import t
from docwen.utils import ocr_utils

pytestmark = pytest.mark.unit
fitz = pytest.importorskip("fitz")


def test_invoice_cn_pdf_respects_original_file_stem_override(tmp_path: Path) -> None:
    pdf_path = tmp_path / "input.pdf"

    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((72, 72), "发票代码：123456789012\n发票号码：12345678\n开票日期：2026年1月2日\n价税合计：100.00")
    doc.save(str(pdf_path))
    doc.close()

    out_dir = tmp_path / "out"
    out_dir.mkdir()

    result = convert_invoice_cn_layout_to_md(
        file_path=str(pdf_path),
        actual_format="pdf",
        output_dir=str(out_dir),
        basename_for_output="b",
        original_file_stem="原始文件名",
    )

    md_path = Path(result["md_path"])
    assert md_path.exists()
    md_text = md_path.read_text(encoding="utf-8")
    assert "aliases:" in md_text
    assert "原始文件名" in md_text


def test_invoice_cn_pdf_multi_page_generates_md_per_page(tmp_path: Path) -> None:
    pdf_path = tmp_path / "input.pdf"
    doc = fitz.open()
    p1 = doc.new_page()
    p1.insert_text((72, 72), "发票号码：11111111\n开票日期：2026年01月01日\n价税合计：1.00")
    p2 = doc.new_page()
    p2.insert_text((72, 72), "发票号码：22222222\n开票日期：2026年01月02日\n价税合计：2.00")
    doc.save(str(pdf_path))
    doc.close()

    out_dir = tmp_path / "out"
    out_dir.mkdir()

    result = convert_invoice_cn_layout_to_md(
        file_path=str(pdf_path),
        actual_format="pdf",
        output_dir=str(out_dir),
        basename_for_output="b_20260101_000000_fromPdf",
        original_file_stem="原始文件名",
    )

    assert result["page_count"] == 2
    md_paths = result["md_paths"]
    assert len(md_paths) == 2

    md1 = Path(md_paths[0])
    md2 = Path(md_paths[1])
    assert md1.exists()
    assert md2.exists()

    section1 = t("conversion.filenames.page_n", n=1)
    section2 = t("conversion.filenames.page_n", n=2)
    assert section1 in md1.stem
    assert section2 in md2.stem

    text1 = md1.read_text(encoding="utf-8")
    text2 = md2.read_text(encoding="utf-8")
    assert "11111111" in text1
    assert "22222222" in text2


def test_invoice_cn_pdf_scanpage_falls_back_to_ocr(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    pdf_path = tmp_path / "scan.pdf"
    doc = fitz.open()
    doc.new_page()
    doc.save(str(pdf_path))
    doc.close()

    out_dir = tmp_path / "out"
    out_dir.mkdir()

    monkeypatch.setattr(
        ocr_utils,
        "extract_text_simple",
        lambda *_a, **_k: "发票号码：33333333\n开票日期：2026年1月2日\n价税合计：3.00",
        raising=True,
    )

    result = convert_invoice_cn_layout_to_md(
        file_path=str(pdf_path),
        actual_format="pdf",
        output_dir=str(out_dir),
        basename_for_output="b",
        original_file_stem="原始文件名",
    )

    md_text = Path(result["md_path"]).read_text(encoding="utf-8")
    assert "33333333" in md_text


def test_invoice_cn_pdf_scanpage_ocr_normalizes_common_misreads(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    pdf_path = tmp_path / "scan.pdf"
    doc = fitz.open()
    doc.new_page()
    doc.save(str(pdf_path))
    doc.close()

    out_dir = tmp_path / "out"
    out_dir.mkdir()

    monkeypatch.setattr(
        ocr_utils,
        "extract_text_simple",
        lambda *_a, **_k: (
            "发票代码：1O3456789O12\n"
            "发票号码：26O47OO000OOO3924O2353\n"
            "校验码：123456789O123456789O\n"
            "开票日期：2O26年O2月11日\n"
            "价税合计：¥256.4O"
        ),
        raising=True,
    )

    result = convert_invoice_cn_layout_to_md(
        file_path=str(pdf_path),
        actual_format="pdf",
        output_dir=str(out_dir),
        basename_for_output="b",
        original_file_stem="原始文件名",
    )

    md_text = Path(result["md_path"]).read_text(encoding="utf-8")
    assert "发票代码: '103456789012'" in md_text
    assert "发票号码: '2604700000000392402353'" in md_text
    assert "校验码: '12345678901234567890'" in md_text
    assert "价税合计: '256.40'" in md_text
