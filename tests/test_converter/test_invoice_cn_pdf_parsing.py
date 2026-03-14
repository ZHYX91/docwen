"""converter 单元测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen.converter.layout2md import invoice_cn
from docwen.converter.layout2md.invoice_cn import convert_invoice_cn_layout_to_md

pytestmark = pytest.mark.unit


def test_invoice_cn_pdf_parses_long_invoice_number_and_rows(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    pdf_path = tmp_path / "invoice.pdf"
    fitz = pytest.importorskip("fitz")
    doc = fitz.open()
    doc.new_page()
    doc.save(str(pdf_path))
    doc.close()

    text = (
        "电子发票（普通发票）\n"
        "发票号码：99990000123456789012\n"
        "开票日期：2026年02月11日\n"
        "购\n买\n方\n信\n息\n名称:张三\n统一社会信用代码/纳税人识别号:911111111111111111\n"
        "销\n售\n方\n信\n息\n名称:示例商贸有限公司\n统一社会信用代码/纳税人识别号:922222222222222222\n"
        "收款账号：6222000000000000000 开户行：中国工商银行示例支行\n"
        "项目名称 规格型号 单 位 数 量 单 价 金 额 税率/征收率 税 额\n"
        "*商品类A*示例商品甲 1×0.8kg 盒 1 104.6 104.60 免税 ***\n"
        "*商品类B*示例商品乙 续行描述 1×1盒 盒 1 106.11 106.11 13% 13.79\n"
        "*商品类C*示例商品丙 1.5kg 1×1盒 盒 1 29.9 29.90 免税 ***\n"
        "*服务类*配送服务费 1 1.77 1.77 13% 0.23\n"
        "合     计 ¥242.38 ¥14.02\n"
        "价税合计（大写） 贰佰伍拾陆圆肆角整 （小写）¥256.40\n"
    )

    spans = []
    header_y = 100.0
    for x0, s in [
        (72.0, "项目名称"),
        (220.0, "规格型号"),
        (300.0, "单 位"),
        (340.0, "数 量"),
        (380.0, "单 价"),
        (430.0, "金 额"),
        (480.0, "税率/征收率"),
        (545.0, "税 额"),
    ]:
        spans.append((x0, header_y, x0 + 10.0, header_y + 10.0, s))

    def add_row(
        y0: float,
        name: str,
        spec: str,
        unit: str,
        qty: str,
        unit_price: str,
        amount: str,
        tax_rate: str,
        tax_amount: str,
    ) -> None:
        spans.append((72.0, y0, 200.0, y0 + 10.0, name))
        spans.append((220.0, y0, 280.0, y0 + 10.0, spec))
        spans.append((300.0, y0, 320.0, y0 + 10.0, unit))
        spans.append((340.0, y0, 360.0, y0 + 10.0, qty))
        spans.append((380.0, y0, 410.0, y0 + 10.0, unit_price))
        spans.append((430.0, y0, 470.0, y0 + 10.0, amount))
        spans.append((480.0, y0, 520.0, y0 + 10.0, tax_rate))
        spans.append((545.0, y0, 580.0, y0 + 10.0, tax_amount))

    add_row(120.0, "*商品类A*示例商品甲", "1×0.8kg", "盒", "1", "104.6", "104.60", "免税", "***")
    add_row(140.0, "*商品类B*示例商品乙 续行描述", "1×1盒", "盒", "1", "106.11", "106.11", "13%", "13.79")
    add_row(160.0, "*商品类C*示例商品丙 1.5kg", "1×1盒", "盒", "1", "29.9", "29.90", "免税", "***")
    add_row(180.0, "*服务类*配送服务费", "", "", "1", "1.77", "1.77", "13%", "0.23")

    spans.append((72.0, 220.0, 90.0, 230.0, "合计"))

    monkeypatch.setattr(
        "docwen.converter.layout2md.invoice_cn._read_pdf_text_and_spans",
        lambda _p: (text, spans),
        raising=True,
    )
    monkeypatch.setattr(
        "docwen.converter.layout2md.invoice_cn._is_scanpage",
        lambda _t: False,
        raising=True,
    )

    out_dir = tmp_path / "out"
    out_dir.mkdir()

    result = convert_invoice_cn_layout_to_md(
        file_path=str(pdf_path),
        actual_format="pdf",
        output_dir=str(out_dir),
        basename_for_output="inv",
        original_file_stem="99990000123456789012",
    )

    md_text = Path(result["md_path"]).read_text(encoding="utf-8")
    assert "99990000123456789012" in md_text
    assert "优化类型:" not in md_text
    assert "发票种类: 电子发票（普通发票）" in md_text
    assert "购买方名称:" in md_text
    assert "张三" in md_text
    assert "购买方纳税人识别号: '911111111111111111'" in md_text
    assert "销售方名称:" in md_text
    assert "示例商贸有限公司" in md_text
    assert "销售方纳税人识别号: '922222222222222222'" in md_text
    assert "购买方地址电话: ''" in md_text
    assert "销售方开户行及账号:" in md_text
    assert "中国工商银行示例支行" in md_text
    assert "6222000000000000000" in md_text
    assert "| *商品类A*示例商品甲 |" in md_text
    assert "| *服务类*配送服务费 |" in md_text


def test_invoice_cn_pdf_spans_handles_split_heji_and_footer(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    pdf_path = tmp_path / "inv.pdf"
    fitz = pytest.importorskip("fitz")
    doc = fitz.open()
    doc.new_page()
    doc.save(str(pdf_path))
    doc.close()

    spans: list[tuple[float, float, float, float, str]] = []

    def add(x0: float, y0: float, text: str) -> None:
        spans.append((x0, y0, x0 + 10, y0 + 10, text))

    add(10.0, 10.0, "发票号码：")
    add(120.0, 10.0, "99990000123456789012")
    add(10.0, 15.0, "开票日期：")
    add(120.0, 15.0, "2026年03月01日")

    add(10.0, 30.0, "名称：")
    add(60.0, 30.0, "示例购买方")
    add(300.0, 30.0, "名称：")
    add(350.0, 30.0, "示例销售方")
    add(10.0, 35.0, "统一社会信用代码/纳税人识别号：")
    add(160.0, 35.0, "911111111111111111")
    add(300.0, 35.0, "统一社会信用代码/纳税人识别号：")
    add(450.0, 35.0, "922222222222222222")

    add(45.0, 50.0, "项目名称")
    add(120.0, 50.0, "规格型号")
    add(190.0, 50.0, "单 位")
    add(260.0, 50.0, "数 量")
    add(335.0, 50.0, "单 价")
    add(405.0, 50.0, "金 额")
    add(445.0, 50.0, "税率/征收率")
    add(550.0, 50.0, "税 额")

    add(12.0, 60.4, "*服务类*示例服务")
    add(190.0, 60.8, "项")
    add(286.0, 60.0, "1")
    add(316.0, 60.0, "100.00")
    add(388.0, 60.0, "100.00")
    add(467.0, 60.0, "6%")
    add(546.0, 60.0, "6.00")
    add(12.0, 72.0, "续行说明")

    add(58.0, 120.0, "合")
    add(103.0, 120.0, "计")
    add(395.0, 119.5, "100.00")
    add(536.0, 119.5, "6.00")
    add(48.0, 130.0, "价税合计（大写）")
    add(447.0, 128.0, "106.00")

    add(55.0, 140.0, "开票人：")
    add(90.0, 140.0, "示例姓名")

    def fake_read(_file_path: str):
        return ("", spans)

    monkeypatch.setattr(invoice_cn, "_read_pdf_text_and_spans", fake_read)
    monkeypatch.setattr(invoice_cn, "_is_scanpage", lambda _t: False)

    out_dir = tmp_path / "out"
    result = convert_invoice_cn_layout_to_md(
        file_path=str(pdf_path),
        actual_format="pdf",
        output_dir=str(out_dir),
        basename_for_output="inv",
        original_file_stem="inv",
    )

    md_text = Path(result["md_path"]).read_text(encoding="utf-8")
    assert "发票号码: '99990000123456789012'" in md_text
    assert "开票日期: 2026年03月01日" in md_text
    assert "购买方名称: 示例购买方" in md_text
    assert "销售方名称: 示例销售方" in md_text
    assert "购买方纳税人识别号: '911111111111111111'" in md_text
    assert "销售方纳税人识别号: '922222222222222222'" in md_text
    assert "价税合计: '106.00'" in md_text
    assert "开票人: 示例姓名" in md_text
    assert "| *服务类*示例服务续行说明 |" in md_text


def test_invoice_cn_pdf_spans_footer_detects_heji_with_punctuation(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    pdf_path = tmp_path / "inv.pdf"
    fitz = pytest.importorskip("fitz")
    doc = fitz.open()
    doc.new_page()
    doc.save(str(pdf_path))
    doc.close()

    spans: list[tuple[float, float, float, float, str]] = []

    def add(x0: float, y0: float, text: str) -> None:
        spans.append((x0, y0, x0 + 10, y0 + 10, text))

    add(45.0, 50.0, "项目名称")
    add(120.0, 50.0, "规格型号")
    add(190.0, 50.0, "单 位")
    add(260.0, 50.0, "数 量")
    add(335.0, 50.0, "单 价")
    add(405.0, 50.0, "金 额")
    add(445.0, 50.0, "税率/征收率")
    add(550.0, 50.0, "税 额")

    add(12.0, 60.0, "*服务类*示例服务")
    add(190.0, 60.0, "项")
    add(286.0, 60.0, "1")
    add(316.0, 60.0, "100.00")
    add(388.0, 60.0, "100.00")
    add(467.0, 60.0, "6%")
    add(546.0, 60.0, "6.00")

    add(55.0, 120.0, "合计：")
    add(395.0, 119.5, "100.00")
    add(536.0, 119.5, "6.00")

    monkeypatch.setattr(
        "docwen.converter.layout2md.invoice_cn._read_pdf_text_and_spans",
        lambda _p: ("", spans),
        raising=True,
    )
    monkeypatch.setattr(
        "docwen.converter.layout2md.invoice_cn._is_scanpage",
        lambda _t: False,
        raising=True,
    )

    out_dir = tmp_path / "out"
    result = convert_invoice_cn_layout_to_md(
        file_path=str(pdf_path),
        actual_format="pdf",
        output_dir=str(out_dir),
        basename_for_output="inv",
        original_file_stem="inv",
    )

    md_text = Path(result["md_path"]).read_text(encoding="utf-8")
    row_line = next(line for line in md_text.splitlines() if line.strip().startswith("| *服务类*示例服务 |"))
    cells = [c.strip() for c in row_line.split("|")[1:-1]]
    assert cells[0] == "*服务类*示例服务"
    assert "合计" not in cells[0]
    assert cells[5] == "100.00"
    assert cells[7] == "6.00"
