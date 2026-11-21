"""Read structured invoice attachments and positioned OFD pages from the package."""

from __future__ import annotations

import posixpath
import re
import xml.etree.ElementTree as etree
import zipfile
from collections.abc import Iterator
from contextlib import suppress
from pathlib import Path
from typing import TYPE_CHECKING

from docwen_plugin_optimizer_invoice_cn.invoice_cn.metadata import (
    merge_page_metadata,
    parse_invoice_metadata_from_text_and_spans,
)
from docwen_plugin_optimizer_invoice_cn.invoice_cn.span_fields import report_metadata_issue
from docwen_plugin_optimizer_invoice_cn.invoice_cn.yaml_schema import TABLE_HEADERS

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import ProgressSink

_Span = tuple[float, float, float, float, str]
_POINTS_PER_MM = 72 / 25.4


def _xml_local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _children(node: etree.Element, local_name: str) -> list[etree.Element]:
    return [child for child in node if _xml_local_name(child.tag) == local_name]


def _xml_values(node: etree.Element, path: tuple[str, ...]) -> list[str]:
    parts = [node]
    for name in path:
        parts = [child for part in parts for child in _children(part, name)]
    return [value for part in parts if (value := (part.text or "").strip())]


def _xml_text(node: etree.Element, *path: str) -> str:
    return next(iter(_xml_values(node, path)), "")


def _member_path(owner: str, reference: str) -> str:
    """Resolve an OFD location inside the ZIP, relative to its owning part."""
    if not reference:
        return ""
    reference = reference.strip().replace("\\", "/")
    path = posixpath.normpath(
        reference.lstrip("/") if reference.startswith("/") else posixpath.join(posixpath.dirname(owner), reference)
    )
    return "" if path in {".", ".."} or path.startswith("../") else path


def _read_xml(package: zipfile.ZipFile, members: dict[str, zipfile.ZipInfo], path: str) -> etree.Element | None:
    member = members.get(path)
    if member is None:
        return None
    try:
        return etree.fromstring(package.read(member))
    except (etree.ParseError, OSError, RuntimeError, NotImplementedError, zipfile.BadZipFile):
        return None


def _invoice_attachments(
    package: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    documents: list[tuple[str, etree.Element]],
) -> tuple[list[etree.Element], bool]:
    attachment_paths: list[str] = []
    broken_relation = False
    has_attachment_list = False
    for doc_path, document in documents:
        relation = _xml_text(document, "Attachments")
        if not relation:
            continue
        has_attachment_list = True
        list_path = _member_path(doc_path, relation)
        attachments = _read_xml(package, members, list_path)
        if attachments is None or _xml_local_name(attachments.tag) != "Attachments":
            broken_relation = True
            continue
        for attachment in _children(attachments, "Attachment"):
            reference = _xml_text(attachment, "FileLoc")
            if reference.lower().endswith(".xml") or (attachment.get("Format") or "").lower() == "xml":
                attachment_paths.append(_member_path(list_path, reference))

    # Producers of the older InvoiceData shape sometimes omit its attachment
    # listing. Restrict that compatibility path to the established basename.
    if not has_attachment_list:
        attachment_paths = [path for path in members if posixpath.basename(path).lower() == "invoicedata.xml"]

    invoices = []
    for path in dict.fromkeys(attachment_paths):
        invoice = _read_xml(package, members, path)
        if invoice is None:
            broken_relation = True
        elif _xml_local_name(invoice.tag) == "Invoice" or (
            has_attachment_list and _xml_local_name(invoice.tag) == "eInvoice"
        ):
            invoices.append(invoice)
    return invoices, broken_relation


def _parse_invoice_xml(
    root: etree.Element, *, sink: ProgressSink | None, location: str
) -> tuple[dict[str, str | None], list[dict[str, str]]]:
    def read_field(node: etree.Element, field: str, paths: list[tuple[str, ...]], *, row: int = 0) -> str:
        values = list(dict.fromkeys(value for path in paths for value in _xml_values(node, path)))
        if len(values) > 1:
            field_location = f"{location}:row-{row}:{field}" if row else f"{location}:{field}"
            report_metadata_issue(
                sink,
                code="INVOICE-ROW-AMBIGUOUS" if row else "INVOICE-METADATA-CONFLICT",
                location=field_location,
                message=f"The invoice attachment contains conflicting values for {field}.",
            )
            return ""
        return next(iter(values), "")

    metadata: dict[str, str | None] = {"优化类型": "invoice_cn"}
    fields = {
        "发票种类": ("InvoiceType",),
        "发票代码": ("InvoiceCode",),
        "发票号码": ("InvoiceNo", "InvoiceNumber"),
        "开票日期": ("IssueDate",),
        "校验码": ("InvoiceCheckCode", "CheckCode"),
        "金额": ("TaxExclusiveTotalAmount", "TotalAmount"),
        "税额": ("TaxTotalAmount", "TotalTax"),
        "价税合计": ("TaxInclusiveTotalAmount", "AmountWithTax"),
        "备注": ("Note", "Remark"),
        "开票人": ("InvoiceClerk", "Issuer"),
        "收款人": ("Payee",),
        "复核": ("Checker", "Reviewer"),
    }
    for key, aliases in fields.items():
        metadata[key] = read_field(root, key, [(alias,) for alias in aliases])

    for role, side in (("购买方", "Buyer"), ("销售方", "Seller")):
        for label, name in (
            ("名称", "Name"),
            ("纳税人识别号", "TaxID"),
            ("地址电话", "AddrTel"),
            ("开户行及账号", "FinancialAccount"),
        ):
            tag = side + name
            field = role + label
            metadata[field] = read_field(root, field, [(side, tag), (tag,)])

    line_fields = (
        ("Item", "GoodsName", "ItemName"),
        ("Specification", "SpecModel"),
        ("MeasurementDimension", "Unit"),
        ("Quantity",),
        ("Price", "UnitPrice"),
        ("Amount",),
        ("TaxScheme", "TaxRate"),
        ("TaxAmount",),
    )
    nodes = (
        [line for infos in _children(root, "GoodsInfos") for line in _children(infos, "GoodsInfo")]
        if _xml_local_name(root.tag) == "eInvoice"
        else [node for node in root.iter() if _xml_local_name(node.tag) == "InvoiceLineInfo"]
    )
    rows = [
        {
            key: read_field(node, key, [(alias,) for alias in aliases], row=row_index)
            for key, aliases in zip(TABLE_HEADERS, line_fields, strict=True)
        }
        for row_index, node in enumerate(nodes, start=1)
    ]
    return metadata, rows


def _positioned_spans(root: etree.Element) -> list[_Span]:
    spans: list[_Span] = []
    for obj in root.iter():
        if _xml_local_name(obj.tag) != "TextObject":
            continue
        try:
            x0, y0, *_ = (float(value) for value in (obj.get("Boundary") or "").split())
            font_size = float(obj.get("Size") or "3")
        except ValueError:
            continue
        for code in _children(obj, "TextCode"):
            text = "".join(code.itertext()).strip()
            if not text:
                continue
            try:
                x = x0 + float(code.get("X") or "0")
                baseline = y0 + float(code.get("Y") or "0")
            except ValueError:
                continue
            # The TextObject boundary is frequently the full cell. Estimate
            # this text run's end from the font size and authored advances.
            width = font_size * sum(1 if ord(char) > 255 else 0.55 for char in text)
            delta_x = code.get("DeltaX") or ""
            if delta_x:
                length = len(text)
                with suppress(ValueError):
                    advances = re.sub(
                        r"g\s+(\d+)\s+(\S+)",
                        lambda match, max_count=length: (match[2] + " ") * min(int(match[1]), max_count),
                        delta_x,
                    )
                    width = sum(float(advance) for advance in advances.split()) + font_size
            spans.append(
                (
                    x * _POINTS_PER_MM,
                    baseline * _POINTS_PER_MM,
                    (x + width) * _POINTS_PER_MM,
                    baseline * _POINTS_PER_MM,
                    text,
                )
            )
    return spans


def _page_spans(
    package: zipfile.ZipFile,
    members: dict[str, zipfile.ZipInfo],
    documents: list[tuple[str, etree.Element]],
) -> Iterator[list[_Span]]:
    if not documents:
        paths = sorted(
            path
            for path in members
            if "/pages/" in f"/{path.casefold()}" and posixpath.basename(path).casefold() == "content.xml"
        )
        for path in paths:
            page = _read_xml(package, members, path)
            yield _positioned_spans(page) if page is not None else []
        return

    for doc_path, document in documents:
        templates = {
            node.get("ID"): _member_path(doc_path, node.get("BaseLoc", ""))
            for common in _children(document, "CommonData")
            for node in _children(common, "TemplatePage")
        }
        for pages in _children(document, "Pages"):
            for page_ref in _children(pages, "Page"):
                path = _member_path(doc_path, page_ref.get("BaseLoc", ""))
                page = _read_xml(package, members, path)
                if page is None:
                    yield []
                    continue
                spans = _positioned_spans(page)
                for source in _children(page, "Template"):
                    template_path = templates.get(source.get("TemplateID"), "")
                    template = _read_xml(package, members, template_path)
                    if template is not None:
                        spans.extend(_positioned_spans(template))
                yield spans


def parse_ofd_invoice(
    file_path: str, *, diagnostic_sink: ProgressSink | None = None
) -> tuple[dict[str, str | None], list[dict[str, str]]]:
    """Use the invoice attachment when available, otherwise parse each physical page."""
    location = Path(file_path).name
    try:
        package = zipfile.ZipFile(file_path)
    except (OSError, zipfile.BadZipFile):
        raise ValueError("The OFD invoice package could not be opened.") from None
    with package:
        members: dict[str, zipfile.ZipInfo] = {}
        duplicates: set[str] = set()
        for member in package.infolist():
            if member.is_dir():
                continue
            path = member.filename.lstrip("/")
            if path in members:
                duplicates.add(path)
            members[path] = member
        if duplicates:
            for path in duplicates:
                del members[path]
            report_metadata_issue(
                diagnostic_sink,
                code="INVOICE-OFD-DUPLICATE-MEMBER",
                location=location,
                message="The OFD package contains ambiguous duplicate parts.",
            )
            if "OFD.xml" in duplicates:
                raise ValueError("The OFD invoice package manifest is ambiguous.")
        manifest = _read_xml(package, members, "OFD.xml")
        documents: list[tuple[str, etree.Element]] = []
        if manifest is not None:
            for body in _children(manifest, "DocBody"):
                path = _member_path("OFD.xml", _xml_text(body, "DocRoot"))
                document = _read_xml(package, members, path)
                if document is not None:
                    documents.append((path, document))
                else:
                    report_metadata_issue(
                        diagnostic_sink,
                        code="INVOICE-OFD-DOCUMENT",
                        location=location,
                        message="An OFD document referenced by the package could not be read.",
                    )

        invoices, broken_relation = _invoice_attachments(package, members, documents)
        if broken_relation or len(invoices) > 1:
            report_metadata_issue(
                diagnostic_sink,
                code="INVOICE-OFD-ATTACHMENT",
                location=location,
                message="The OFD invoice attachment is unreadable or the package contains multiple invoice attachments.",
            )
        if len(invoices) == 1:
            return _parse_invoice_xml(invoices[0], sink=diagnostic_sink, location=location)
        if len(invoices) > 1:
            raise ValueError("The OFD package contains multiple invoice attachments.")

        metadata: dict[str, str | None] = {}
        rows: list[dict[str, str]] = []
        field_pages: dict[str, int] = {}
        any_spans = False
        for page_number, spans in enumerate(_page_spans(package, members, documents), start=1):
            page_location = f"{location}:page-{page_number}"
            if not spans:
                report_metadata_issue(
                    diagnostic_sink,
                    code="INVOICE-OFD-PAGE-TEXT",
                    location=page_location,
                    message="No positioned text could be read from this OFD page.",
                )
                continue
            any_spans = True
            text = "\n".join(span[-1] for span in sorted(spans, key=lambda span: (round(span[1]), span[0])))
            page_metadata, page_rows = parse_invoice_metadata_from_text_and_spans(
                text, spans, diagnostic_sink=diagnostic_sink, location=page_location
            )
            merge_page_metadata(metadata, field_pages, page_metadata, page_number, diagnostic_sink, page_location)
            rows.extend(page_rows)
        if not any_spans:
            raise ValueError("未找到电子发票 XML，且无法从 OFD 页面提取文本")
        return metadata, rows
