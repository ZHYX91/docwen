"""Report incomplete and unreconciled invoices without changing source values."""

from __future__ import annotations

import re
from decimal import Decimal, localcontext
from typing import TYPE_CHECKING

from docwen_plugin_optimizer_invoice_cn.invoice_cn.ocr_normalize import (
    _normalize_ocr_amount,
    _normalize_positioned_amount,
)
from docwen_plugin_optimizer_invoice_cn.invoice_cn.rows import _EXEMPT_TAX, _TAX_RATE, _label_text
from docwen_plugin_optimizer_invoice_cn.invoice_cn.span_fields import report_metadata_issue

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import ProgressSink

_MONEY_FIELDS = ("金额", "税额", "价税合计")
_REQUIRED_FIELDS = ("发票号码", "开票日期", "购买方名称", "销售方名称", "销售方纳税人识别号")


def _amount(value: str | None, *, ocr: bool) -> Decimal | None:
    if not value or not re.search(r"\d", value):
        return None
    normalized = (_normalize_ocr_amount if ocr else _normalize_positioned_amount)(value)
    return Decimal(normalized) if normalized else None


def validate_invoice(
    metadata: dict[str, str | None],
    rows: list[dict[str, str]],
    *,
    sink: ProgressSink | None,
    location: str,
    ocr: bool = False,
) -> None:
    summary = {field: _amount(metadata.get(field), ocr=ocr) for field in _MONEY_FIELDS}
    missing = [field for field in _REQUIRED_FIELDS if not metadata.get(field)]
    missing.extend(field for field, value in summary.items() if value is None)
    if missing or not rows:
        fields = ", ".join(missing + ([] if rows else ["商品明细"]))
        report_metadata_issue(
            sink,
            code="INVOICE-INCOMPLETE",
            location=location,
            message=f"Missing or unreadable invoice fields: {fields}.",
        )

    if metadata.get("发票号码") and metadata["发票号码"] == metadata.get("校验码"):
        report_metadata_issue(
            sink,
            code="INVOICE-IDENTITY-CONFLICT",
            location=location,
            message="The source invoice number and check code are identical.",
        )

    detail = {"金额": [], "税额": []}
    incomplete_rows = []
    for index, row in enumerate(rows, start=1):
        rate = _label_text(row.get("税率", ""))
        raw_tax = row.get("税额", "")
        tax = _amount(raw_tax, ocr=ocr)
        if tax is None and rate in {"免税", "不征税"} and _EXEMPT_TAX.fullmatch(raw_tax):
            tax = Decimal(0)
        amount = _amount(row.get("金额"), ocr=ocr)
        detail["金额"].append(amount)
        detail["税额"].append(tax)
        if (
            amount is None
            or tax is None
            or not row.get("商品名称")
            or not (_TAX_RATE.fullmatch(rate) or re.fullmatch(r"0(?:\.\d+)?|1(?:\.0+)?", rate))
        ):
            incomplete_rows.append(str(index))
    if incomplete_rows:
        report_metadata_issue(
            sink,
            code="INVOICE-ROW-INCOMPLETE",
            location=location,
            message=f"Required detail fields are missing or unreadable on row(s) {', '.join(incomplete_rows)}.",
        )

    # Derive enough precision to add the actual lexical amounts, including
    # inputs with long integer parts or more than two fractional places.
    amounts = [value for value in (*summary.values(), *detail["金额"], *detail["税额"]) if value is not None]
    if not amounts:
        return
    max_integer_digits = max(value.adjusted() + 1 for value in amounts)
    max_fraction_digits = max(-int(value.as_tuple().exponent) for value in amounts)
    mismatch = []
    with localcontext() as context:
        context.prec = max(28, max_integer_digits + max_fraction_digits + len(str(len(amounts))) + 1)
        subtotal, tax_total, grand_total = (summary[field] for field in _MONEY_FIELDS)
        if (
            subtotal is not None
            and tax_total is not None
            and grand_total is not None
            and subtotal + tax_total != grand_total
        ):
            mismatch.append("金额 + 税额 ≠ 价税合计")
        for field, entries in detail.items():
            stated = summary[field]
            if (
                stated is not None
                and entries
                and all(value is not None for value in entries)
                and sum(value for value in entries if value is not None) != stated
            ):
                mismatch.append(f"明细{field}合计 ≠ 票面{field}")
    if mismatch:
        report_metadata_issue(
            sink,
            code="INVOICE-TOTAL-MISMATCH",
            location=location,
            message="Invoice amounts do not reconcile: " + ", ".join(mismatch) + ".",
        )
