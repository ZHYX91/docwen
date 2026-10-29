"""Shared amount parsing retains structured and positioned OFD field signs."""

from __future__ import annotations

from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

import pytest

from tests.integration.test_invoice_native_field_identity import _page_items, _run

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]
_OFD = "http://www.ofdspec.org/2016"


@pytest.mark.parametrize("structured", [False, True], ids=["positioned-content", "invoice-data"])
@pytest.mark.parametrize(
    "amount,tax,total,quantity", [("-100.00", "-13.00", "-113.00", "-1"), ("100.00", "13.00", "113.00", "1")]
)
def test_new_ofd_file_keeps_signed_metadata_and_detail_fields(
    tmp_path: Path, structured: bool, amount: str, tax: str, total: str, quantity: str
) -> None:
    items, expected_row = _page_items(amount=amount, tax=tax, total=total, quantity=quantity)
    source = tmp_path / "signed-fields.ofd"
    page = ET.Element(f"{{{_OFD}}}Page")
    area = ET.SubElement(page, f"{{{_OFD}}}Area")
    ET.SubElement(area, f"{{{_OFD}}}PhysicalBox").text = "0 0 1000 600"
    content = ET.SubElement(page, f"{{{_OFD}}}Content")
    layer = ET.SubElement(content, f"{{{_OFD}}}Layer", ID="10002")
    for index, (x, y, text) in enumerate(items, start=2):
        node = ET.SubElement(
            layer, f"{{{_OFD}}}TextObject", ID=str(index), Boundary=f"{x} {y} 200 10", Font="10000", Size="8"
        )
        ET.SubElement(node, f"{{{_OFD}}}TextCode", X="0", Y="0").text = text
    invoice = ET.Element("Invoice")
    for key, value in {
        "InvoiceNumber": "87654321",
        "BuyerName": "购买单位",
        "SellerName": "销售单位",
        "TotalAmount": amount,
        "TotalTax": tax,
        "AmountWithTax": total,
    }.items():
        ET.SubElement(invoice, key).text = value
    line = ET.SubElement(invoice, "InvoiceLineInfo")
    for key, value in zip(
        ("GoodsName", "SpecModel", "Unit", "Quantity", "UnitPrice", "Amount", "TaxRate", "TaxAmount"),
        expected_row,
        strict=True,
    ):
        ET.SubElement(line, key).text = value
    with ZipFile(source, "w", compression=ZIP_DEFLATED) as package:
        package.writestr(
            "OFD.xml",
            f'<ofd:OFD xmlns:ofd="{_OFD}" Version="1.0" DocType="OFD"><ofd:DocBody><ofd:DocInfo><ofd:DocID>0123456789abcdef0123456789abcdef</ofd:DocID></ofd:DocInfo><ofd:DocRoot>Doc_0/Document.xml</ofd:DocRoot></ofd:DocBody></ofd:OFD>',
        )
        package.writestr(
            "Doc_0/Document.xml",
            f'<ofd:Document xmlns:ofd="{_OFD}"><ofd:CommonData><ofd:MaxUnitID>10002</ofd:MaxUnitID><ofd:PageArea><ofd:PhysicalBox>0 0 1000 600</ofd:PhysicalBox></ofd:PageArea><ofd:PublicRes>PublicRes.xml</ofd:PublicRes></ofd:CommonData><ofd:Pages><ofd:Page ID="10001" BaseLoc="Pages/Page_0/Content.xml"/></ofd:Pages></ofd:Document>',
        )
        package.writestr(
            "Doc_0/PublicRes.xml",
            f'<ofd:Res xmlns:ofd="{_OFD}" BaseLoc="."><ofd:Fonts><ofd:Font ID="10000" FontName="SimSun" FamilyName="宋体"/></ofd:Fonts></ofd:Res>',
        )
        package.writestr("Doc_0/Pages/Page_0/Content.xml", ET.tostring(page, encoding="utf-8", xml_declaration=True))
        if structured:
            package.writestr(
                "Doc_0/Attachs/InvoiceData.xml", ET.tostring(invoice, encoding="utf-8", xml_declaration=True)
            )
    with ZipFile(source) as package:
        source_page = ET.fromstring(package.read("Doc_0/Pages/Page_0/Content.xml"))
        source_document = ET.fromstring(package.read("Doc_0/Document.xml"))
        source_resources = ET.fromstring(package.read("Doc_0/PublicRes.xml"))
        ids = [
            int(node.attrib["ID"])
            for root in (source_page, source_document, source_resources)
            for node in root.iter()
            if "ID" in node.attrib
        ]
        assert len(ids) == len(set(ids))
        assert max(ids) == int(source_document.findtext(f"{{{_OFD}}}CommonData/{{{_OFD}}}MaxUnitID", "0"))
        source_font = source_resources.find(f"{{{_OFD}}}Fonts/{{{_OFD}}}Font")
        assert source_font is not None
        assert source_font.get("ID") == "10000"
        assert {node.get("Font") for node in source_page.iter(f"{{{_OFD}}}TextObject")} == {"10000"}
        assert [node.text for node in source_page.findall(f".//{{{_OFD}}}TextCode")] == [text for _x, _y, text in items]
        if structured:
            source_invoice = ET.fromstring(package.read("Doc_0/Attachs/InvoiceData.xml"))
            assert [source_invoice.findtext(field) for field in ("TotalAmount", "TotalTax", "AmountWithTax")] == [
                amount,
                tax,
                total,
            ]
            source_line = source_invoice.find("InvoiceLineInfo")
            assert source_line is not None
            assert [node.text for node in source_line] == expected_row
    metadata, rows, _ = _run(source, tmp_path)
    assert {field: metadata[field] for field in ("金额", "税额", "价税合计")} == {
        "金额": amount,
        "税额": tax,
        "价税合计": total,
    }
    assert all(isinstance(metadata[field], str) for field in ("金额", "税额", "价税合计"))
    assert rows == [expected_row]
