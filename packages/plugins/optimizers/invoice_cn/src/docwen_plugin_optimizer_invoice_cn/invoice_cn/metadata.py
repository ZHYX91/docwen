"""Invoice metadata extraction from compact text and PDF spans."""

from __future__ import annotations

import re

from docwen_plugin_optimizer_invoice_cn.invoice_cn.ocr_normalize import (
    _compact_text,
    _normalize_ocr_amount,
    _normalize_ocr_date,
    _normalize_ocr_digits,
    _normalize_ocr_tax_id,
    _regex_first,
)


def parse_invoice_metadata_from_compact_text(
    compact: str,
) -> dict[str, str | None]:
    """Extract invoice metadata fields from compacted text via regex.

    Args:
        compact: Whitespace-free text from the invoice.

    Returns:
        A dict of field name → value.  Missing fields are ``None``.
    """
    metadata: dict[str, str | None] = {"优化类型": "invoice_cn"}

    # ── Basic fields ─────────────────────────────────────────────────
    invoice_kind = _regex_first(
        compact,
        [
            r"(电子发票（普通发票）|电子发票（增值税专用发票）|增值税专用发票|增值税普通发票|电子普通发票)",
        ],
    )
    invoice_code_raw = _regex_first(compact, [r"发票代码[:：]?([0-9OolIZzsB]{10,12})"])
    invoice_number_raw = _regex_first(compact, [r"发票号码[:：]?([0-9OolIZzsB]{8,30})"])
    issue_date = _regex_first(
        compact,
        [
            r"开票日期[:：]?([0-9OolIZzsB]{4}[年./·•-][0-9OolIZzsB]{1,2}[月./·•-][0-9OolIZzsB]{1,2}日?)",
        ],
    )
    issue_date = _normalize_ocr_date(issue_date)
    invoice_code = _normalize_ocr_digits(invoice_code_raw)
    invoice_number = _normalize_ocr_digits(invoice_number_raw)

    # ── Amount / tax ─────────────────────────────────────────────────
    amount_with_tax = _regex_first(
        compact,
        [
            r"价税合[计汁].*?小写.*?[¥￥]?([0-9OolIZzsB]+(?:\.[0-9OolIZzsB]{1,2})?)",
            r"价税合[计汁][:：]?[¥￥]?([0-9OolIZzsB]+(?:\.[0-9OolIZzsB]{1,2})?)",
        ],
    )
    amount_with_tax = _normalize_ocr_amount(amount_with_tax)

    total_amount_match = re.search(
        r"合计[¥￥]?([0-9OolIZzsB]+(?:\.[0-9OolIZzsB]{1,2})?)[¥￥]?([0-9OolIZzsB]+(?:\.[0-9OolIZzsB]{1,2})?)",
        compact,
    )

    # ── Buyer / seller names ─────────────────────────────────────────
    buyer_name = _regex_first(compact, [r"购买方信息名称[:：]?(.+?)(?=统一社会信用代码|销售方信息)"])
    seller_name = _regex_first(compact, [r"销售方信息名称[:：]?(.+?)(?=统一社会信用代码|项目名称|货物或应税劳务)"])

    # ── Buyer / seller blocks for tax IDs and banks ──────────────────
    buyer_block = ""
    seller_block = ""
    bi = compact.find("购买方信息")
    si = compact.find("销售方信息")
    if bi != -1:
        end = si if (si != -1 and si > bi) else len(compact)
        buyer_block = compact[bi:end]
    if si != -1:
        end_candidates: list[int] = []
        for m in ("项目名称", "货物或应税劳务"):
            idx = compact.find(m, si + 1)
            if idx != -1:
                end_candidates.append(idx)
        end = min(end_candidates) if end_candidates else len(compact)
        seller_block = compact[si:end]

    buyer_tax_id = _regex_first(
        buyer_block,
        [
            r"统一社会信用代码/纳税人识别号[:：]?([0-9A-Z]{15,20})",
            r"统一社会信用代码[:：]?([0-9A-Z]{15,20})",
            r"纳税人识别号[:：]?([0-9A-Z]{15,20})",
        ],
    )
    seller_tax_id = _regex_first(
        seller_block,
        [
            r"统一社会信用代码/纳税人识别号[:：]?([0-9A-Z]{15,20})",
            r"统一社会信用代码[:：]?([0-9A-Z]{15,20})",
            r"纳税人识别号[:：]?([0-9A-Z]{15,20})",
        ],
    )
    buyer_tax_id = _normalize_ocr_tax_id(buyer_tax_id)
    seller_tax_id = _normalize_ocr_tax_id(seller_tax_id)

    # ── Check code ───────────────────────────────────────────────────
    check_code_raw = _regex_first(compact, [r"校验码[:：]?([0-9OolIZzsB]{20})"])
    check_code = _normalize_ocr_digits(check_code_raw)

    # ── Populate metadata ────────────────────────────────────────────
    if invoice_kind:
        metadata["发票种类"] = invoice_kind
    if invoice_code:
        metadata["发票代码"] = invoice_code
    if invoice_number:
        metadata["发票号码"] = invoice_number

    # Fallback date search
    if not issue_date:
        issue_date = _normalize_ocr_date(
            _regex_first(
                compact,
                [
                    r"((?:20[0-9OolIZzsB]{2})[年./·•-][0-9OolIZzsB]{1,2}[月./·•-][0-9OolIZzsB]{1,2}日?)",
                ],
            )
        )
    if issue_date:
        metadata["开票日期"] = issue_date
    if check_code:
        metadata["校验码"] = check_code
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

    # ── Fallback digit-search for code / number / check-code ─────────
    if not invoice_code or not invoice_number or not check_code:
        candidates_raw = re.findall(r"[0-9OolIZzsB]{8,30}", compact)
        candidates: list[str] = []
        for raw in candidates_raw:
            d = _normalize_ocr_digits(raw)
            if d and len(d) >= 8:
                candidates.append(d)

        date_digits = ""
        if issue_date:
            m = re.search(r"(20[0-9]{2})年([0-9]{2})月([0-9]{2})日", issue_date)
            if m:
                date_digits = f"{m.group(1)}{m.group(2)}{m.group(3)}"
        if not issue_date:
            for d in candidates:
                m = re.search(r"(20[0-9]{2})([0-1][0-9])([0-3][0-9])", d)
                if not m:
                    continue
                yyyymmdd = f"{m.group(1)}{m.group(2)}{m.group(3)}"
                issue_date_val = f"{m.group(1)}年{int(m.group(2)):02d}月{int(m.group(3)):02d}日"
                metadata["开票日期"] = issue_date_val
                date_digits = yyyymmdd
                break

        if not check_code:
            for d in candidates:
                if len(d) == 20:
                    metadata["校验码"] = d
                    break

        if not invoice_code:
            for d in candidates:
                if len(d) in {10, 11, 12}:
                    metadata["发票代码"] = d
                    break

        if not invoice_number:
            for d in candidates:
                if len(d) == 8 and d != date_digits:
                    metadata["发票号码"] = d
                    break
        if not invoice_number and date_digits:
            for d in candidates:
                if date_digits in d:
                    rest = d.replace(date_digits, "")
                    if len(rest) == 8:
                        metadata["发票号码"] = rest
                        break

    return metadata


def parse_invoice_metadata_from_pdf_spans(
    spans: list[tuple[float, float, float, float, str]],
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

    def find_right_value(*, label_contains: str, value_re: str, y_tol: float = 2.5) -> str:
        for y, x0, _x1, t in items:
            if label_contains not in t:
                continue
            candidates: list[tuple[float, float, str]] = []
            for y2, x02, _x12, t2 in items:
                if abs(y2 - y) > y_tol:
                    continue
                if x02 <= x0 + 5:
                    continue
                m = re.fullmatch(value_re, _compact_text(t2))
                if not m:
                    continue
                candidates.append((abs(y2 - y), x02, m.group(0)))
            if candidates:
                candidates.sort()
                return candidates[0][2]
        return ""

    def find_right_text(*, label_contains: str, y_tol: float = 2.5) -> str:
        for y, x0, _x1, t in items:
            if label_contains not in t:
                continue
            candidates: list[tuple[float, str]] = []
            for y2, x02, _x12, t2 in items:
                if abs(y2 - y) > y_tol:
                    continue
                if x02 <= x0 + 5:
                    continue
                tc = _compact_text(t2)
                if not tc:
                    continue
                if tc.endswith(("：", ":")):
                    continue
                if label_contains in tc:
                    continue
                candidates.append((x02, t2.strip()))
            if candidates:
                candidates.sort()
                return candidates[0][1]
        return ""

    invoice_number = find_right_value(label_contains="发票号码", value_re=r"[0-9]{8,30}")
    issue_date = find_right_value(
        label_contains="开票日期",
        value_re=r"20[0-9]{2}年[0-9]{1,2}月[0-9]{1,2}日",
    )
    issuer = find_right_text(label_contains="开票人")
    reviewer = find_right_text(label_contains="复核")
    payee = find_right_text(label_contains="收款人")

    # Detect left/right split for buyer vs seller columns
    name_labels = [(y, x0, x1) for (y, x0, x1, t) in items if _compact_text(t) in {"名称：", "名称:"}]
    split_x: float | None = None
    if len(name_labels) >= 2:
        for i in range(len(name_labels) - 1):
            y0, x0, _x1 = name_labels[i]
            y1, x1, _x2 = name_labels[i + 1]
            if abs(y0 - y1) <= 2.5:
                split_x = (x0 + x1) / 2.0
                break

    if split_x is None:
        tax_label_x = sorted(
            {
                x0
                for _y, x0, _x1, t in items
                if "统一社会信用代码" in _compact_text(t) or "纳税人识别号" in _compact_text(t)
            }
        )
        if len(tax_label_x) >= 3:
            gaps = [(tax_label_x[index + 1] - tax_label_x[index], index) for index in range(len(tax_label_x) - 1)]
            largest_gap, gap_index = max(gaps)
            if largest_gap >= 50.0:
                split_x = (tax_label_x[gap_index] + tax_label_x[gap_index + 1]) / 2.0

    def pick_name_after_label(label_y: float, label_x0: float, *, side_left: bool) -> str:
        candidates: list[tuple[float, float, str]] = []
        for y, x0, _x1, t in items:
            if abs(y - label_y) <= 2.5 and x0 > label_x0 + 5:
                if split_x is not None and (x0 < split_x) != side_left:
                    continue
                tc = _compact_text(t)
                if re.fullmatch(r"[0-9A-Z]{15,20}", tc) or re.fullmatch(r"[0-9]{8,30}", tc):
                    continue
                if tc in {"名称：", "名称:"} or "统一社会信用代码" in tc or "纳税人识别号" in tc:
                    continue
                candidates.append((0.0, x0, t))
        if candidates:
            candidates.sort()
            return candidates[0][2].strip()

        for y, x0, _x1, t in items:
            if not (label_y + 1.0 <= y <= label_y + 30.0):
                continue
            if split_x is not None and (x0 < split_x) != side_left:
                continue
            if x0 < label_x0 + 5:
                continue
            tc = _compact_text(t)
            if re.fullmatch(r"[0-9A-Z]{15,20}", tc) or re.fullmatch(r"[0-9]{8,30}", tc):
                continue
            if tc in {"名称：", "名称:"} or "统一社会信用代码" in tc or "纳税人识别号" in tc:
                continue
            candidates.append((y - label_y, x0, t))
        if candidates:
            candidates.sort()
            return candidates[0][2].strip()
        return ""

    buyer_name = ""
    seller_name = ""
    if split_x is not None:
        for y, x0, _x1 in name_labels:
            if x0 < split_x and not buyer_name:
                buyer_name = pick_name_after_label(y, x0, side_left=True)
            if x0 >= split_x and not seller_name:
                seller_name = pick_name_after_label(y, x0, side_left=False)

    def pick_tax_id(*, side_left: bool) -> str:
        def is_requested_side(x0: float) -> bool:
            return split_x is None or (x0 < split_x) == side_left

        for y, x0, _x1, t in items:
            tc = _compact_text(t)
            if "统一社会信用代码" not in tc and "纳税人识别号" not in tc:
                continue
            if not is_requested_side(x0):
                continue
            tail = ""
            if "：" in t or ":" in t:
                tail = re.split(r"[:：]", t, maxsplit=1)[-1]
            tail = _compact_text(tail).upper()
            if re.fullmatch(r"[0-9A-Z]{15,20}", tail or ""):
                return tail
            for y2, _x02, _x12, t2 in items:
                if abs(y2 - y) > 2.5:
                    continue
                if _x02 <= x0 + 5:
                    continue
                if not is_requested_side(_x02):
                    continue
                v = _compact_text(t2).upper()
                if re.fullmatch(r"[0-9A-Z]{15,20}", v):
                    return v
            for y2, _x02, _x12, t2 in items:
                if not (y + 1.0 <= y2 <= y + 25.0):
                    continue
                if not is_requested_side(_x02):
                    continue
                v = _compact_text(t2).upper()
                if re.fullmatch(r"[0-9A-Z]{15,20}", v):
                    return v
        return ""

    buyer_tax_id = pick_tax_id(side_left=True)
    seller_tax_id = pick_tax_id(side_left=False)

    # ── 价税合计 (total with tax) ────────────────────────────────────
    amount_with_tax = ""
    total_label = [t for t in items if "价税合计" in _compact_text(t[3])]
    if total_label:
        total_label.sort(key=lambda t: (t[0], t[1]))
        y0, x0, _x1, _t = total_label[0]
        money = []
        for y, x, _x1, s in items:
            compact_value = _compact_text(s)
            if abs(y - y0) > 30 or x <= x0:
                continue
            if not re.fullmatch(r"[¥￥]?-?[0-9]+(?:\.[0-9]{1,2})?", compact_value):
                continue
            money.append((abs(y - y0), x, _normalize_ocr_amount(compact_value)))
        if money:
            money.sort()
            amount_with_tax = money[0][2]

    # ── Amount / tax (金额 / 税额) via y-clustering ───────────────────
    money_2dec = [(y, x, s) for y, x, _x1, s in items if re.fullmatch(r"[0-9]+\.[0-9]{2}", _compact_text(s))]
    amount = ""
    tax = ""
    if total_label:
        y_total = total_label[0][0]
        buckets: dict[float, list[str]] = {}
        for y, _x, s in money_2dec:
            if y >= y_total - 10:
                continue
            if y < y_total - 120:
                continue
            key = round(y, 1)
            buckets.setdefault(key, []).append(s)
        if buckets:
            best_y = max(buckets.keys())
            vals = buckets[best_y]
            if len(vals) >= 2:
                a = max(vals, key=lambda v: float(v))
                b = min(vals, key=lambda v: float(v))
                amount = a
                tax = b
            elif len(vals) == 1:
                amount = vals[0]
                tax = ""

    out: dict[str, str] = {}
    if invoice_number:
        out["发票号码"] = invoice_number
    if issue_date:
        out["开票日期"] = issue_date
    if buyer_name:
        out["购买方名称"] = buyer_name
    if seller_name:
        out["销售方名称"] = seller_name
    if buyer_tax_id:
        out["购买方纳税人识别号"] = buyer_tax_id
    if seller_tax_id:
        out["销售方纳税人识别号"] = seller_tax_id
    if amount:
        out["金额"] = amount
    if tax:
        out["税额"] = tax
    if amount_with_tax:
        out["价税合计"] = amount_with_tax
    if issuer:
        out["开票人"] = issuer
    if reviewer:
        out["复核"] = reviewer
    if payee:
        out["收款人"] = payee
    return out


def parse_invoice_metadata_from_text_and_spans(
    text: str,
    spans: list[tuple[float, float, float, float, str]],
) -> tuple[dict[str, str | None], list[dict[str, str]]]:
    """Extract metadata from both compact text (regex) and PDF spans (positional).

    This is the main entry point for PDF invoice metadata extraction.
    Span-based extraction overrides text-based extraction for fields
    that both sources detect.

    Args:
        text: Raw text from PDF page.
        spans: List of (x0, y0, x1, y1, text) tuples from PyMuPDF.

    Returns:
        A tuple of (metadata_dict, rows_list).
        Rows may be empty if row parsing fails.
    """
    compact = _compact_text(text)
    metadata = parse_invoice_metadata_from_compact_text(compact)

    span_meta = parse_invoice_metadata_from_pdf_spans(spans)
    for k, v in span_meta.items():
        if v:
            metadata[k] = v

    from docwen_plugin_optimizer_invoice_cn.invoice_cn.rows import (
        parse_invoice_rows_from_pdf_spans,
        parse_invoice_rows_from_pdf_text,
    )

    rows = parse_invoice_rows_from_pdf_spans(spans=spans)
    if not rows:
        rows = parse_invoice_rows_from_pdf_text(text, prefer_marked=False)

    return metadata, rows
