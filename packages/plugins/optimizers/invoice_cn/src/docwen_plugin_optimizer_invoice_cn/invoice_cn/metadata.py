"""Invoice metadata extraction from compact text and PDF spans."""

from __future__ import annotations

import re
from typing import TYPE_CHECKING
from unicodedata import normalize

from docwen_plugin_optimizer_invoice_cn.invoice_cn.ocr_normalize import (
    AMOUNT_TOKEN_PATTERN,
    _compact_text,
    _normalize_ocr_amount,
    _normalize_ocr_date,
    _normalize_ocr_digits,
    _regex_first,
)
from docwen_plugin_optimizer_invoice_cn.invoice_cn.span_fields import (
    amount_metadata,
    party_metadata,
    report_metadata_issue,
)

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import ProgressSink

# Whitespace was removed by the compact-text adapter. Only the end of the
# text or an explicit following field label can delimit this amount. Unknown
# suffixes (including malformed grouping) remain unowned, not numeric prefixes.
_MONEY_END = r"(?=$|价税合[计汁]|收款人|复核|开票人|备注|购买方|销售方|发票|开票日期|校验码|机器编号)"
_DIGITS = r"[0-9OolIZzsB]"
_DATE = rf"{_DIGITS}{{4}}[年./·•-]{_DIGITS}{{1,2}}[月./·•-]{_DIGITS}{{1,2}}日?"
_LABELLED_PATTERNS = {
    "发票代码": (rf"(?:{_DIGITS}{{20}}|{_DIGITS}{{10,12}})(?!{_DIGITS})", _normalize_ocr_digits),
    "发票号码": (rf"{_DIGITS}{{8,30}}(?!{_DIGITS})", _normalize_ocr_digits),
    "开票日期": (_DATE, _normalize_ocr_date),
    "校验码": (rf"{_DIGITS}{{20}}(?!{_DIGITS})", _normalize_ocr_digits),
}


def _unique_compact_field(block: str, pattern: str) -> str:
    """A role-labelled text block owns a field only when its value is unique."""
    values = {match.group(1).strip() for match in re.finditer(pattern, block)}
    return next(iter(values)) if len(values) == 1 else ""


def parse_invoice_metadata_from_compact_text(
    compact: str,
) -> dict[str, str | None]:
    """Extract invoice metadata fields from compacted text via regex.

    Args:
        compact: Whitespace-free text from the invoice.

    Returns:
        A dict of field name → value.  Missing fields are ``None``.
    """
    compact = normalize("NFKC", compact)
    metadata: dict[str, str | None] = {"优化类型": "invoice_cn"}

    # ── Basic fields ─────────────────────────────────────────────────
    invoice_kind = _regex_first(
        compact,
        [
            r"(电子发票\((?:普通发票|增值税专用发票)\)|增值税专用发票|增值税普通发票|电子普通发票)",
        ],
    )
    for field, (pattern, normalize_value) in _LABELLED_PATTERNS.items():
        raw = _unique_compact_field(compact, rf"{field}:?({pattern})")
        if raw and (value := normalize_value(raw)):
            metadata[field] = value

    # ── Amount / tax ─────────────────────────────────────────────────
    amount_with_tax = _regex_first(
        compact,
        [
            rf"价税合[计汁](?:[（(]大写[）)][^()（）]*?)?[（(]?小写[）):：]*({AMOUNT_TOKEN_PATTERN}){_MONEY_END}",
            rf"价税合[计汁][:：]?({AMOUNT_TOKEN_PATTERN}){_MONEY_END}",
        ],
    )
    amount_with_tax = _normalize_ocr_amount(amount_with_tax)

    total_amount_match = re.search(
        rf"(?<!价税)合计[:：]?({AMOUNT_TOKEN_PATTERN})(?=[+\-−－＋]?[¥￥])({AMOUNT_TOKEN_PATTERN}){_MONEY_END}",
        compact,
    )

    # ── Buyer / seller blocks for tax IDs and banks ──────────────────
    buyer_block = ""
    seller_block = ""
    bi = compact.find("购买方信息")
    si = compact.find("销售方信息")
    if bi != -1:
        end = si if (si != -1 and si > bi) else len(compact)
        buyer_block = compact[bi:end]
    if si != -1:
        end_candidates: list[int] = [bi] if bi > si else []
        for m in ("项目名称", "货物或应税劳务"):
            idx = compact.find(m, si + 1)
            if idx != -1:
                end_candidates.append(idx)
        end = min(end_candidates) if end_candidates else len(compact)
        seller_block = compact[si:end]

    tax_pattern = r"(?:统一社会信用代码/纳税人识别号|统一社会信用代码|纳税人识别号):?([0-9A-Z]{15,20})(?![0-9A-Z])"
    name_pattern = r"名称:?([^:]+?)(?=统一社会信用代码|纳税人识别号|开户行及账号|名称:|$)"
    buyer_name = _unique_compact_field(buyer_block, name_pattern)
    seller_name = _unique_compact_field(seller_block, name_pattern)
    buyer_tax_id = _unique_compact_field(buyer_block, tax_pattern)
    seller_tax_id = _unique_compact_field(seller_block, tax_pattern)

    # ── Populate metadata ────────────────────────────────────────────
    if invoice_kind:
        metadata["发票种类"] = invoice_kind.replace("(", "（").replace(")", "）")
    if amount_with_tax:
        metadata["价税合计"] = amount_with_tax
    if total_amount_match:
        amount = _normalize_ocr_amount((total_amount_match.group(1) or "").strip())
        tax = _normalize_ocr_amount((total_amount_match.group(2) or "").strip())
        if amount:
            metadata["金额"] = amount
        if tax:
            metadata["税额"] = tax
    if buyer_name:
        metadata["购买方名称"] = buyer_name
    if buyer_tax_id:
        metadata["购买方纳税人识别号"] = buyer_tax_id
    if seller_name:
        metadata["销售方名称"] = seller_name
    if seller_tax_id:
        metadata["销售方纳税人识别号"] = seller_tax_id

    # ── Bank accounts ────────────────────────────────────────────────
    buyer_bank = _regex_first(
        buyer_block,
        [r"开户行及账号[:：]?(.+?)(?=销售方信息|项目名称|货物或应税劳务|$)"],
    )
    seller_bank = _regex_first(
        seller_block,
        [r"开户行及账号[:：]?(.+?)(?=项目名称|货物或应税劳务|$)"],
    )
    if buyer_bank:
        metadata["购买方开户行及账号"] = buyer_bank
    if seller_bank:
        metadata["销售方开户行及账号"] = seller_bank

    collect_account = _regex_first(
        compact,
        [
            r"收款账号[:：]?([0-9]{10,30})",
            r"收款帐号[:：]?([0-9]{10,30})",
        ],
    )
    collect_bank = _regex_first(
        compact,
        [
            r"开户行[:：]?(.+?)(?=20[0-9]{2}年|开票日期|校验码|发票号码|$)",
        ],
    )
    if collect_bank or collect_account:
        v = " ".join([x for x in [collect_bank, collect_account] if x])
        if v:
            metadata.setdefault("销售方开户行及账号", v)

    return metadata


def parse_invoice_metadata_from_pdf_spans(
    spans: list[tuple[float, float, float, float, str]],
    *,
    diagnostic_sink: ProgressSink | None = None,
    location: str = "",
) -> dict[str, str]:
    """Extract invoice metadata from PDF span positions.

    Uses positional heuristics: finds label spans and looks for values
    to the right on the same or nearby lines.

    Args:
        spans: List of (x0, y0, x1, y1, text) tuples from PyMuPDF.

    Returns:
        A dict of field name → value (only populated fields, never ``None``).
    """
    items: list[tuple[float, float, float, str]] = []
    for x0, y0, x1, _y1, s in spans:
        t = (s or "").strip()
        if not t:
            continue
        items.append((float(y0), float(x0), float(x1), t))
    items.sort(key=lambda t: (t[0], t[1]))

    from docwen_plugin_optimizer_invoice_cn.invoice_cn.span_fields import labelled_span_values, remark_metadata

    parties = party_metadata(items, sink=diagnostic_sink, location=location)
    amounts = amount_metadata(items, spans, sink=diagnostic_sink, location=location)

    out = labelled_span_values(items, _LABELLED_PATTERNS, sink=diagnostic_sink, location=location)
    out.update(parties)
    out.update(amounts)
    if remark := remark_metadata(items):
        out["备注"] = remark
    return out


def parse_invoice_metadata_from_text_and_spans(
    text: str,
    spans: list[tuple[float, float, float, float, str]],
    *,
    diagnostic_sink: ProgressSink | None = None,
    location: str = "",
) -> tuple[dict[str, str | None], list[dict[str, str]]]:
    """Extract metadata from both compact text (regex) and PDF spans (positional).

    This is the main entry point for PDF invoice metadata extraction.
    A complete role-labelled text value retains ownership. Positional
    evidence fills missing values and reports disagreements without silently
    replacing the stronger text fact. Callers provide one physical page.

    Args:
        text: Raw text from PDF page.
        spans: List of (x0, y0, x1, y1, text) tuples from PyMuPDF.

    Returns:
        A tuple of (metadata_dict, rows_list).
        Rows may be empty if row parsing fails.
    """
    compact = _compact_text(text)
    metadata = parse_invoice_metadata_from_compact_text(compact)
    if spans:
        # The compact adapter removes whitespace before OCR matching, which can
        # erase the boundary between two native amount values inside a span.
        for field in ("金额", "税额", "价税合计"):
            metadata.pop(field, None)
        # Flattening a column's vertical labels can associate its name with
        # the other column. Text owns a party only with a complete role label.
        for role in ("购买方", "销售方"):
            if f"{role}信息" not in normalize("NFKC", text):
                for suffix in ("名称", "纳税人识别号", "开户行及账号"):
                    metadata.pop(role + suffix, None)

    span_meta = parse_invoice_metadata_from_pdf_spans(spans, diagnostic_sink=diagnostic_sink, location=location)
    for key, value in span_meta.items():
        if not value:
            continue
        previous = metadata.get(key)
        if previous and normalize("NFKC", previous) != normalize("NFKC", value):
            report_metadata_issue(
                diagnostic_sink,
                code="INVOICE-METADATA-CONFLICT",
                location=f"{location}:{key}",
                message=f"Conflicting {key} text and position evidence within this page; kept the complete role-labelled text value.",
            )
        elif not previous:
            metadata[key] = value

    from docwen_plugin_optimizer_invoice_cn.invoice_cn.rows import (
        parse_invoice_rows_from_pdf_spans,
        parse_invoice_rows_from_pdf_text,
    )

    rows = parse_invoice_rows_from_pdf_spans(spans=spans, diagnostic_sink=diagnostic_sink, location=location)
    if not rows:
        rows = parse_invoice_rows_from_pdf_text(text, prefer_marked=False)

    return metadata, rows


def merge_page_metadata(
    metadata: dict[str, str | None],
    field_pages: dict[str, int],
    page_metadata: dict[str, str | None],
    page_number: int,
    diagnostic_sink: ProgressSink | None,
    location: str,
) -> None:
    for field, value in page_metadata.items():
        if not value:
            continue
        previous = metadata.get(field)
        if previous is None:
            metadata[field] = value
            field_pages[field] = page_number
        elif normalize("NFKC", previous) != normalize("NFKC", value):
            report_metadata_issue(
                diagnostic_sink,
                code="INVOICE-METADATA-CONFLICT",
                location=f"{location}:{field}",
                message=f"Conflicting {field} on page-{field_pages[field]} and page-{page_number}; kept the first page-owned value.",
            )
