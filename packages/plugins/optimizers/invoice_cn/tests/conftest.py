"""Shared fixtures for Invoice plugin tests."""

from __future__ import annotations

import zipfile
from io import BytesIO
from pathlib import Path

import pytest

# Try to find a CJK-capable font for PDF generation.
_CJK_FONT_PATH: str | None = None
for _candidate in (
    "C:/Windows/Fonts/msyh.ttc",
    "C:/Windows/Fonts/simsun.ttc",
    "C:/Windows/Fonts/STSONG.TTF",
):
    if Path(_candidate).exists():
        _CJK_FONT_PATH = _candidate
        break


def _create_test_pdf_invoice(path: Path, text: str | None = None) -> None:
    """Create a minimal PDF file with invoice-like text content.

    If *text* is provided, it is written into the PDF page.
    Otherwise a default simulated invoice text is used.

    Uses a CJK-capable font when available so that PyMuPDF text
    extraction preserves Chinese characters.
    """
    import fitz

    if text is None:
        text = (
            "电子发票（普通发票）\n"
            "发票代码：123456789012\n"
            "发票号码：87654321\n"
            "开票日期：2024年01月15日\n"
            "校验码：12345678901234567890\n"
            "购买方信息\n"
            "名称：测试购买方公司\n"
            "统一社会信用代码/纳税人识别号：91110108MA01ABCD1X\n"
            "销售方信息\n"
            "名称：测试销售方公司\n"
            "统一社会信用代码/纳税人识别号：91110108MA01EFGH2Y\n"
            "项目名称 规格型号 单位 数量 单价 金额 税率 税额\n"
            "*测试商品一 件 2 100.00 200.00 13% 26.00\n"
            "*测试商品二 箱 1 500.00 500.00 13% 65.00\n"
            "合计 ¥700.00 ¥91.00\n"
            "价税合计 ¥791.00\n"
            "收款人：张三  复核：李四  开票人：王五\n"
        )

    doc = fitz.open()
    page = doc.new_page(width=595, height=842)

    if _CJK_FONT_PATH:
        try:
            page.insert_text((72, 72), text, fontfile=_CJK_FONT_PATH, fontsize=10)
        except Exception:
            page.insert_text((72, 72), text, fontsize=10)
    else:
        page.insert_text((72, 72), text, fontsize=10)

    doc.save(str(path))
    doc.close()


def _create_test_ofd_invoice(path: Path, *, with_xml: bool = True) -> None:
    """Create a minimal OFD file (ZIP container) for testing.

    When *with_xml* is True, the ZIP contains an InvoiceData.xml.
    When False, the ZIP contains content.xml pages instead (fallback mode).
    """
    if with_xml:
        xml_content = """<?xml version="1.0" encoding="UTF-8"?>
<Invoice>
    <InvoiceCode>123456789012</InvoiceCode>
    <InvoiceNumber>87654321</InvoiceNumber>
    <IssueDate>2024年01月15日</IssueDate>
    <BuyerName>测试购买方公司</BuyerName>
    <SellerName>测试销售方公司</SellerName>
    <TotalAmount>700.00</TotalAmount>
    <TotalTax>91.00</TotalTax>
    <AmountWithTax>791.00</AmountWithTax>
    <InvoiceLineInfo>
        <GoodsName>测试商品一</GoodsName>
        <SpecModel></SpecModel>
        <Unit>件</Unit>
        <Quantity>2</Quantity>
        <UnitPrice>100.00</UnitPrice>
        <Amount>200.00</Amount>
        <TaxRate>13%</TaxRate>
        <TaxAmount>26.00</TaxAmount>
    </InvoiceLineInfo>
    <InvoiceLineInfo>
        <GoodsName>测试商品二</GoodsName>
        <SpecModel></SpecModel>
        <Unit>箱</Unit>
        <Quantity>1</Quantity>
        <UnitPrice>500.00</UnitPrice>
        <Amount>500.00</Amount>
        <TaxRate>13%</TaxRate>
        <TaxAmount>65.00</TaxAmount>
    </InvoiceLineInfo>
</Invoice>"""
        buf = BytesIO()
        with zipfile.ZipFile(buf, "w") as zf:
            zf.writestr("InvoiceData.xml", xml_content)
        path.write_bytes(buf.getvalue())
    else:
        # Fallback mode: OFD with content.xml containing TextObject elements
        content_xml = """<?xml version="1.0" encoding="UTF-8"?>
<Page>
    <TextObject Boundary="10 10 500 800">
        <TextCode X="10" Y="10">电子发票（普通发票）</TextCode>
        <TextCode X="10" Y="30">发票号码：87654321</TextCode>
        <TextCode X="10" Y="50">开票日期：2024年01月15日</TextCode>
        <TextCode X="10" Y="70">名称：测试购买方公司</TextCode>
        <TextCode X="10" Y="90">名称：测试销售方公司</TextCode>
        <TextCode X="10" Y="110">项目名称 规格型号 单位 数量 单价 金额 税率 税额</TextCode>
        <TextCode X="10" Y="130">*测试商品一 件 2 100.00 200.00 13% 26.00</TextCode>
        <TextCode X="10" Y="150">*测试商品二 箱 1 500.00 500.00 13% 65.00</TextCode>
        <TextCode X="10" Y="170">合计 ¥700.00 ¥91.00</TextCode>
        <TextCode X="10" Y="190">价税合计 ¥791.00</TextCode>
    </TextObject>
</Page>"""
        buf = BytesIO()
        with zipfile.ZipFile(buf, "w") as zf:
            zf.writestr("Pages/Page_0/Content.xml", content_xml)
        path.write_bytes(buf.getvalue())


@pytest.fixture
def sample_invoice_pdf_path(tmp_path: Path) -> Path:
    """A single-page PDF with simulated invoice text content."""
    path = tmp_path / "invoice.pdf"
    _create_test_pdf_invoice(path)
    return path


@pytest.fixture
def sample_invoice_ofd_path(tmp_path: Path) -> Path:
    """An OFD file with InvoiceData.xml containing invoice data."""
    path = tmp_path / "invoice.ofd"
    _create_test_ofd_invoice(path, with_xml=True)
    return path


@pytest.fixture
def sample_invoice_ofd_fallback_path(tmp_path: Path) -> Path:
    """An OFD file without InvoiceData.xml, using content.xml fallback."""
    path = tmp_path / "invoice_fallback.ofd"
    _create_test_ofd_invoice(path, with_xml=False)
    return path


@pytest.fixture
def sample_simple_pdf_path(tmp_path: Path) -> Path:
    """A minimal single-page PDF with basic text (for error handling tests)."""
    import fitz

    path = tmp_path / "simple.pdf"
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)
    page.insert_text((72, 72), "Hello World", fontsize=12)
    doc.save(str(path))
    doc.close()
    return path
