"""Synthetic invoice packages and PDF layouts through Markdown publication."""

from __future__ import annotations

from pathlib import Path
from xml.etree import ElementTree
from zipfile import ZipFile

import pytest

from docwen_plugin_optimizer_invoice_cn.manifest import build_manifest
from tests.integration.test_invoice_native_field_identity import _page_items, _pdf, _run

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]
_OFD = "{http://www.ofdspec.org/2016}"


def test_linked_einvoice_attachment_defines_fields_and_all_detail_lines(tmp_path: Path) -> None:
    source = tmp_path / "attached.ofd"
    ns = "{urn:synthetic:einvoice}"
    invoice = ElementTree.Element(ns + "eInvoice")
    for key, value in {
        "InvoiceCode": "00000000000123456789",
        "InvoiceNo": "00000028",
        "InvoiceCheckCode": "00000000000987654321",
        "IssueDate": "2024年10月12日",
        "TaxExclusiveTotalAmount": "99.88",
        "TaxTotalAmount": "5.94",
        "TaxInclusiveTotalAmount": "105.82",
        "InvoiceClerk": "示例经办",
        "Note": "合成备注",
    }.items():
        ElementTree.SubElement(invoice, ns + key).text = value
    for role, tax_id, value in (
        ("Buyer", "123456789012345678", "合成购买方"),
        ("Seller", "987654321098765432", "合成销售方"),
    ):
        party = ElementTree.SubElement(invoice, ns + role)
        ElementTree.SubElement(party, ns + role + "Name").text = value
        ElementTree.SubElement(party, ns + role + "TaxID").text = tax_id

    lines = ElementTree.SubElement(invoice, ns + "GoodsInfos")
    expected = []
    for index in range(100):
        signed = index == 99
        amount, rate, tax = ("-0.11", "免税", "***") if signed else ("1.01", "6%", "0.06")
        values = [
            f"商品{max(index, 1)}",
            "样式A",
            "件",
            "" if signed else "1",
            "" if signed else "1.010000001",
            amount,
            rate,
            tax,
        ]
        line = ElementTree.SubElement(lines, ns + "GoodsInfo")
        for name, value in zip(
            ("Item", "Specification", "MeasurementDimension", "Quantity", "Price", "Amount", "TaxScheme", "TaxAmount"),
            values,
            strict=True,
        ):
            ElementTree.SubElement(line, ns + name).text = value
        expected.append(values)
    with ZipFile(source, "w") as package:
        package.writestr("OFD.xml", "<OFD><DocBody><DocRoot>/body/Document.xml</DocRoot></DocBody></OFD>")
        package.writestr("body/Document.xml", "<Document><Attachments>rels/Attachments.xml</Attachments></Document>")
        package.writestr(
            "body/rels/Attachments.xml",
            "<Attachments><Attachment Format='xml'><FileLoc>../values/data.xml</FileLoc></Attachment></Attachments>",
        )
        package.writestr("body/values/data.xml", ElementTree.tostring(invoice, encoding="utf-8"))
        package.writestr("body/values/InvoiceData.xml", "<Invoice><InvoiceNumber>99999999</InvoiceNumber></Invoice>")

    metadata, rows, result = _run(source, tmp_path)
    assert (metadata["发票代码"], metadata["发票号码"], metadata["校验码"]) == (
        "00000000000123456789",
        "00000028",
        "00000000000987654321",
    )
    assert (metadata["购买方名称"], metadata["销售方名称"], metadata["备注"], metadata["开票人"]) == (
        "合成购买方",
        "合成销售方",
        "合成备注",
        "示例经办",
    )
    assert (metadata["购买方纳税人识别号"], metadata["销售方纳税人识别号"]) == (
        "123456789012345678",
        "987654321098765432",
    )
    assert rows == expected
    assert not any(
        d.code in {"INVOICE-INCOMPLETE", "INVOICE-ROW-INCOMPLETE", "INVOICE-TOTAL-MISMATCH"} for d in result.diagnostics
    )
    [rule] = [rule for rule in build_manifest().capability_rules if rule.source_formats == ("ofd",)]
    assert not rule.required_capabilities


def _write_page(package: ZipFile, path: str, items: list[tuple[float, float, str]], template: str = "") -> None:
    page = ElementTree.Element(_OFD + "Page")
    if template:
        ElementTree.SubElement(page, _OFD + "Template", TemplateID=template)
    layer = ElementTree.SubElement(ElementTree.SubElement(page, _OFD + "Content"), _OFD + "Layer")
    for x, y, text in items:
        text_object = ElementTree.SubElement(layer, _OFD + "TextObject", Boundary=f"{x} {y} 30 8", Size="8")
        ElementTree.SubElement(text_object, _OFD + "TextCode", X="0", Y="0").text = text
    package.writestr(path, ElementTree.tostring(page, encoding="utf-8"))


def test_positioned_ofd_uses_only_the_template_owned_by_each_declared_page(tmp_path: Path) -> None:
    source = tmp_path / "positioned.ofd"
    items, first_row = _page_items(mode="buyer-only", item_name="第一页", include_totals=False)
    other_items, second_row = _page_items(mode="seller-only", item_name="第二页")
    detail_top = 300.0
    header = [(x, y, text) for x, y, text in items if y == detail_top]
    first_values = [(x, y, text) for x, y, text in items if y != detail_top]
    second_values = [(x, y, text) for x, y, text in other_items if y != detail_top]
    with ZipFile(source, "w") as package:
        package.writestr("OFD.xml", "<OFD><DocBody><DocRoot>root/layout.xml</DocRoot></DocBody></OFD>")
        package.writestr(
            "root/layout.xml",
            """<Document><CommonData>
            <TemplatePage ID="1" BaseLoc="docs/style.xml"/><TemplatePage ID="2" BaseLoc="docs/unused.xml"/>
            </CommonData><Pages><Page ID="3" BaseLoc="docs/b.xml"/><Page ID="4" BaseLoc="docs/a.xml"/>
            </Pages></Document>""",
        )
        _write_page(package, "root/docs/a.xml", second_values, template="1")
        _write_page(package, "root/docs/b.xml", first_values, template="1")
        _write_page(package, "root/docs/unused.xml", [(40, 58, "发票号码：11111111")])
        _write_page(package, "root/docs/style.xml", [*header, (10, 430, "备注")])
        package.writestr("original_invoice.xml", "<eInvoice><InvoiceNo>99999999</InvoiceNo></eInvoice>")

    metadata, rows, result = _run(source, tmp_path)
    assert metadata["发票号码"] == "87654321"
    assert (metadata["购买方名称"], metadata["销售方名称"]) == ("购买单位", "销售单位")
    assert rows == [first_row, second_row]
    assert any(d.code == "INVOICE-TOTAL-MISMATCH" for d in result.diagnostics)
    assert not any(d.code == "INVOICE-METADATA-CONFLICT" for d in result.diagnostics)


def test_pdf_keeps_source_totals_and_warns_about_missing_and_inconsistent_fields(tmp_path: Path) -> None:
    source = tmp_path / "unreconciled.pdf"
    items, expected = _page_items(mode="seller-only", amount="12.00", tax="1.56", total="14.00")
    # The 20-digit number has only an invoice-number label. The accompanying
    # unlabelled number cannot become an invoice code or a check code.
    items[1] = (40.0, 58.0, "发票号码：00000000000000000123")
    items.insert(2, (510.0, 58.0, "123456789012"))
    items = [(x, y, "￥11.00" if text == "￥12.00" else text) for x, y, text in items]
    _pdf(source, [items])

    metadata, rows, result = _run(source, tmp_path)
    assert (metadata["发票号码"], metadata["发票代码"], metadata["校验码"]) == ("00000000000000000123", "", "")
    assert metadata["购买方名称"] == ""
    assert (metadata["金额"], metadata["税额"], metadata["价税合计"]) == ("11.00", "1.56", "14.00")
    assert rows == [expected]
    assert {"INVOICE-INCOMPLETE", "INVOICE-TOTAL-MISMATCH"} <= {d.code for d in result.diagnostics}


def test_duplicate_ofd_attachment_path_never_selects_one_invoice_version(tmp_path: Path) -> None:
    source = tmp_path / "duplicates.ofd"
    items, expected_row = _page_items()
    with ZipFile(source, "w") as package:
        package.writestr("OFD.xml", "<OFD><DocBody><DocRoot>Document.xml</DocRoot></DocBody></OFD>")
        package.writestr(
            "Document.xml",
            '<Document><Pages><Page ID="1" BaseLoc="page.xml"/></Pages><Attachments>rels.xml</Attachments></Document>',
        )
        package.writestr(
            "rels.xml", "<Attachments><Attachment Format='xml'><FileLoc>draft.xml</FileLoc></Attachment></Attachments>"
        )
        package.writestr("draft.xml", "<eInvoice><InvoiceNo>11111111</InvoiceNo></eInvoice>")
        package.writestr("/draft.xml", "<eInvoice><InvoiceNo>22222222</InvoiceNo></eInvoice>")
        _write_page(package, "page.xml", items)
    metadata, rows, result = _run(source, tmp_path)
    assert metadata["发票号码"] == "87654321"
    assert rows == [expected_row]
    [warning] = [d for d in result.diagnostics if d.code == "INVOICE-OFD-DUPLICATE-MEMBER"]
    assert warning.location == "duplicates.ofd"
    assert "draft.xml" not in warning.message
