import logging
import re
import tempfile
import threading
import zipfile
from bisect import bisect_right
from pathlib import Path
from typing import Any, cast

import xml.etree.ElementTree as etree

from docwen.translation import t
from docwen.utils.text_utils import format_yaml_value
from docwen.utils.workspace_manager import finalize_output

logger = logging.getLogger(__name__)

INVOICE_CN_YAML_SCHEMA = [
    "发票种类",
    "发票代码",
    "发票号码",
    "开票日期",
    "校验码",
    "购买方名称",
    "购买方纳税人识别号",
    "购买方地址电话",
    "购买方开户行及账号",
    "销售方名称",
    "销售方纳税人识别号",
    "销售方地址电话",
    "销售方开户行及账号",
    "金额",
    "税额",
    "价税合计",
    "备注",
    "收款人",
    "复核",
    "开票人",
]


def _build_yaml_frontmatter(*, file_stem: str, metadata: dict[str, str | None], include_empty: bool = False) -> str:
    lines: list[str] = ["---"]

    safe_stem = format_yaml_value(file_stem)
    lines.append("aliases:")
    lines.append(f"  - {safe_stem}")

    title_key = t("yaml_keys.title", default="标题")
    lines.append(f"{title_key}: {safe_stem}")

    for key, value in metadata.items():
        if value is None:
            if include_empty:
                lines.append(f"{key}: ''")
            continue
        value = str(value).strip()
        if not value:
            if include_empty:
                lines.append(f"{key}: ''")
            continue
        lines.append(f"{key}: {format_yaml_value(value)}")

    lines.append("---")
    lines.append("")
    return "\n".join(lines)


def _escape_table_cell(value: str) -> str:
    return str(value).replace("|", "\\|").replace("\n", " ").strip()


def _render_markdown_table(*, headers: list[str], rows: list[dict[str, str]]) -> str:
    normalized_rows: list[list[str]] = []
    for row in rows:
        normalized_rows.append([_escape_table_cell(row.get(h, "")) for h in headers])

    if not normalized_rows:
        normalized_rows = [[_escape_table_cell("（未识别）")] + [""] * (len(headers) - 1)]

    out: list[str] = []
    out.append("| " + " | ".join(headers) + " |")
    out.append("| " + " | ".join(["---"] * len(headers)) + " |")
    for r in normalized_rows:
        out.append("| " + " | ".join(r) + " |")
    out.append("")
    return "\n".join(out)


def _xml_local_name(tag: str) -> str:
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def _xml_first_text(root: etree.Element, local_name: str) -> str:
    for el in root.iter():
        if _xml_local_name(el.tag) != local_name:
            continue
        v = (el.text or "").strip()
        if v:
            return v
    return ""


def _parse_ofd_invoice(file_path: str) -> tuple[dict[str, str | None], list[dict[str, str]]]:
    xml_bytes = None
    with zipfile.ZipFile(file_path) as zf:
        for name in zf.namelist():
            if name.lower().endswith("invoicedata.xml"):
                xml_bytes = zf.read(name)
                break

    if not xml_bytes:
        items = _extract_ofd_items(file_path)
        if not items:
            raise ValueError("未找到 InvoiceData.xml 且无法从 OFD 内容提取文本")
        text = _extract_ofd_text_from_items(items)
        metadata = _parse_invoice_metadata_from_ofd_items(items)
        rows = _parse_invoice_rows_from_ofd_items(items)
        if not rows:
            rows = _parse_invoice_rows_from_pdf_text(text, prefer_marked=True)
        return metadata, rows

    root = etree.fromstring(xml_bytes)

    metadata: dict[str, str | None] = {
        "优化类型": "invoice_cn",
        "发票代码": _xml_first_text(root, "InvoiceCode"),
        "发票号码": _xml_first_text(root, "InvoiceNumber"),
        "开票日期": _xml_first_text(root, "IssueDate"),
        "购买方名称": _xml_first_text(root, "BuyerName"),
        "销售方名称": _xml_first_text(root, "SellerName"),
        "金额": _xml_first_text(root, "TotalAmount"),
        "税额": _xml_first_text(root, "TotalTax"),
        "价税合计": _xml_first_text(root, "AmountWithTax"),
    }

    rows: list[dict[str, str]] = []
    for node in root.iter():
        if _xml_local_name(node.tag) != "InvoiceLineInfo":
            continue
        row = {
            "商品名称": _xml_first_text(node, "GoodsName") or _xml_first_text(node, "ItemName"),
            "规格型号": _xml_first_text(node, "SpecModel"),
            "单位": _xml_first_text(node, "Unit"),
            "数量": _xml_first_text(node, "Quantity"),
            "单价": _xml_first_text(node, "UnitPrice"),
            "金额": _xml_first_text(node, "Amount"),
            "税率": _xml_first_text(node, "TaxRate"),
            "税额": _xml_first_text(node, "TaxAmount"),
        }
        if any(v.strip() for v in row.values()):
            rows.append(row)

    return metadata, rows


def _regex_first(text: str, patterns: list[str]) -> str:
    for p in patterns:
        m = re.search(p, text, flags=re.MULTILINE)
        if m:
            return (m.group(1) or "").strip()
    return ""


def _normalize_ocr_digits(value: str) -> str:
    s = (value or "").strip()
    if not s:
        return ""
    table = str.maketrans(
        {
            "O": "0",
            "o": "0",
            "I": "1",
            "l": "1",
            "Z": "2",
            "z": "2",
            "S": "5",
            "s": "5",
            "B": "8",
        }
    )
    s = s.translate(table)
    s = re.sub(r"[^0-9]", "", s)
    return s


def _normalize_ocr_amount(value: str) -> str:
    s = (value or "").strip()
    if not s:
        return ""
    table = str.maketrans(
        {
            "O": "0",
            "o": "0",
            "I": "1",
            "l": "1",
            "Z": "2",
            "z": "2",
            "S": "5",
            "s": "5",
            "B": "8",
            "，": ",",
            "．": ".",
            "。": ".",
        }
    )
    s = s.translate(table)
    s = s.replace("¥", "").replace("￥", "").replace(",", "")
    s = re.sub(r"[^0-9.]", "", s)
    if s.count(".") > 1:
        parts = s.split(".")
        s = parts[0] + "." + "".join(parts[1:])
    if "." in s:
        left, right = s.split(".", 1)
        if not right:
            s = left + ".00"
        elif len(right) == 1:
            s = left + "." + right + "0"
    return s


def _normalize_ocr_tax_id(value: str) -> str:
    s = (value or "").strip()
    if not s:
        return ""
    table = str.maketrans({"O": "0", "o": "0", "I": "1", "l": "1", "Z": "2", "z": "2", "S": "5", "s": "5"})
    s = s.translate(table).upper()
    s = re.sub(r"[^0-9A-Z]", "", s)
    return s


def _normalize_ocr_date(value: str) -> str:
    s = (value or "").strip()
    if not s:
        return ""
    table = str.maketrans({"O": "0", "o": "0", "I": "1", "l": "1", "Z": "2", "z": "2", "S": "5", "s": "5", "B": "8"})
    s_norm = s.translate(table)
    m = re.search(r"(20[0-9]{2})[年./·•-]([0-9]{1,2})[月./·•-]([0-9]{1,2})", s_norm)
    if not m:
        return s_norm.strip()
    y = m.group(1)
    mm = int(m.group(2))
    dd = int(m.group(3))
    return f"{y}年{mm:02d}月{dd:02d}日"


def _compact_text(text: str) -> str:
    s = re.sub(r"\s+", "", text or "")
    return s.replace("\ufeff", "").replace("\u200b", "").replace("\u200c", "").replace("\u200d", "")


def _read_pdf_text_and_spans(file_path: str) -> tuple[str, list[tuple[float, float, float, float, str]]]:
    import fitz

    doc = fitz.open(file_path)
    try:
        text_parts: list[str] = []
        spans: list[tuple[float, float, float, float, str]] = []
        for page in doc:
            text_parts.append(str(page.get_text("text")))
            d = cast(dict[str, Any], page.get_text("dict"))
            for block in d.get("blocks") or []:
                if not isinstance(block, dict):
                    continue
                for line in block.get("lines") or []:
                    if not isinstance(line, dict):
                        continue
                    for span in line.get("spans") or []:
                        if not isinstance(span, dict):
                            continue
                        s = (span.get("text") or "").strip()
                        if not s:
                            continue
                        x0, y0, x1, y1 = span.get("bbox") or (0.0, 0.0, 0.0, 0.0)
                        spans.append((float(x0), float(y0), float(x1), float(y1), s))
        return "\n".join(text_parts), spans
    finally:
        doc.close()


MIN_TEXT_LENGTH_FOR_INVOICE = 20


def _is_scanpage(text: str) -> bool:
    return len(_compact_text(text)) < MIN_TEXT_LENGTH_FOR_INVOICE


def _get_pdf_page_count(file_path: str) -> int:
    import fitz

    doc = fitz.open(file_path)
    try:
        return len(doc)
    finally:
        doc.close()


def _render_pdf_page_to_png(*, file_path: str, page_index: int, png_path: str) -> str:
    import fitz

    doc = fitz.open(file_path)
    try:
        page = doc[page_index]
        dpi = 300
        zoom = max(72, int(dpi)) / 72.0
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        pix.save(png_path)
        try:
            from PIL import Image, ImageEnhance, ImageFilter, ImageOps

            img = Image.open(png_path)
            img = ImageOps.exif_transpose(img)
            img = img.convert("L")
            img = ImageOps.autocontrast(img)
            img = img.filter(ImageFilter.UnsharpMask(radius=2, percent=180, threshold=3))
            img = ImageEnhance.Contrast(img).enhance(1.6)

            hist = img.histogram()
            total = sum(hist)
            if total > 0:
                sum_total = 0.0
                for i, h in enumerate(hist):
                    sum_total += float(i * h)
                sum_b = 0.0
                w_b = 0
                max_var = -1.0
                threshold = 160
                for i in range(256):
                    w_b += int(hist[i])
                    if w_b == 0:
                        continue
                    w_f = total - w_b
                    if w_f == 0:
                        break
                    sum_b += float(i * hist[i])
                    m_b = sum_b / float(w_b)
                    m_f = (sum_total - sum_b) / float(w_f)
                    var_between = float(w_b) * float(w_f) * (m_b - m_f) * (m_b - m_f)
                    if var_between > max_var:
                        max_var = var_between
                        threshold = i

                t = max(90, min(210, int(threshold)))
                img = img.point(lambda x: 255 if x >= t else 0, mode="L")
                img.save(png_path, format="PNG", optimize=True)
        except Exception:
            pass
        return png_path
    finally:
        doc.close()


def _read_pdf_text_and_spans_single_page(
    file_path: str, page_index: int
) -> tuple[str, list[tuple[float, float, float, float, str]]]:
    import fitz

    doc = fitz.open(file_path)
    try:
        page = doc[page_index]
        text = str(page.get_text("text"))
        spans: list[tuple[float, float, float, float, str]] = []
        d = cast(dict[str, Any], page.get_text("dict"))
        for block in d.get("blocks") or []:
            if not isinstance(block, dict):
                continue
            for line in block.get("lines") or []:
                if not isinstance(line, dict):
                    continue
                for span in line.get("spans") or []:
                    if not isinstance(span, dict):
                        continue
                    s = (span.get("text") or "").strip()
                    if not s:
                        continue
                    x0, y0, x1, y1 = span.get("bbox") or (0.0, 0.0, 0.0, 0.0)
                    spans.append((float(x0), float(y0), float(x1), float(y1), s))
        return text, spans
    finally:
        doc.close()


def _parse_pdf_invoice_from_text_and_spans(
    text: str, spans: list[tuple[float, float, float, float, str]]
) -> tuple[dict[str, str | None], list[dict[str, str]]]:
    compact = _compact_text(text)

    metadata = _parse_invoice_metadata_from_compact_text(compact)
    span_meta = _parse_invoice_metadata_from_pdf_spans(spans)
    for k, v in span_meta.items():
        if v:
            metadata[k] = v

    rows = _parse_invoice_rows_from_pdf_spans(spans=spans)
    if not rows:
        rows = _parse_invoice_rows_from_pdf_text(text, prefer_marked=False)

    return metadata, rows


def _parse_pdf_invoice_single_page(
    file_path: str, page_index: int
) -> tuple[dict[str, str | None], list[dict[str, str]]]:
    text, spans = _read_pdf_text_and_spans_single_page(file_path, page_index)
    return _parse_pdf_invoice_from_text_and_spans(text, spans)


def _parse_invoice_metadata_from_pdf_spans(spans: list[tuple[float, float, float, float, str]]) -> dict[str, str]:
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

    invoice_number = find_right_value(label_contains="发票号码", value_re=r"[0-9]{8,30}")
    issue_date = find_right_value(label_contains="开票日期", value_re=r"20[0-9]{2}年[0-9]{1,2}月[0-9]{1,2}日")

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

    issuer = find_right_text(label_contains="开票人")
    reviewer = find_right_text(label_contains="复核")
    payee = find_right_text(label_contains="收款人")

    name_labels = [(y, x0, x1) for (y, x0, x1, t) in items if _compact_text(t) in {"名称：", "名称:"}]
    split_x: float | None = None
    if len(name_labels) >= 2:
        for i in range(len(name_labels) - 1):
            y0, x0, _x1 = name_labels[i]
            y1, x1, _x2 = name_labels[i + 1]
            if abs(y0 - y1) <= 2.5:
                split_x = (x0 + x1) / 2.0
                break

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
        for y, x0, x1 in name_labels:
            if x0 < split_x and not buyer_name:
                buyer_name = pick_name_after_label(y, x0, side_left=True)
            if x0 >= split_x and not seller_name:
                seller_name = pick_name_after_label(y, x0, side_left=False)

    def pick_tax_id(*, side_left: bool) -> str:
        for y, x0, _x1, t in items:
            tc = _compact_text(t)
            if "统一社会信用代码" not in tc and "纳税人识别号" not in tc:
                continue
            if split_x is not None and (x0 < split_x) != side_left:
                continue
            tail = ""
            if "：" in t or ":" in t:
                tail = re.split(r"[:：]", t, maxsplit=1)[-1]
            tail = _compact_text(tail).upper()
            if re.fullmatch(r"[0-9A-Z]{15,20}", tail or ""):
                return tail
            for y2, x02, _x12, t2 in items:
                if abs(y2 - y) > 2.5:
                    continue
                if x02 <= x0 + 5:
                    continue
                v = _compact_text(t2).upper()
                if re.fullmatch(r"[0-9A-Z]{15,20}", v):
                    return v
            for y2, x02, _x12, t2 in items:
                if not (y + 1.0 <= y2 <= y + 25.0):
                    continue
                v = _compact_text(t2).upper()
                if re.fullmatch(r"[0-9A-Z]{15,20}", v):
                    return v
        return ""

    buyer_tax_id = pick_tax_id(side_left=True)
    seller_tax_id = pick_tax_id(side_left=False)

    amount_with_tax = ""
    total_label = [t for t in items if "价税合计" in _compact_text(t[3])]
    if total_label:
        total_label.sort(key=lambda t: (t[0], t[1]))
        y0, x0, _x1, _t = total_label[0]
        money = [
            (abs(y - y0), x, s)
            for y, x, _x1, s in items
            if re.fullmatch(r"[0-9]+(?:\.[0-9]{2})?", _compact_text(s)) and abs(y - y0) <= 30 and x > x0
        ]
        if money:
            money.sort()
            amount_with_tax = money[0][2]

    money_2dec = [(y, x, s) for y, x, _x1, s in items if re.fullmatch(r"[0-9]+\.[0-9]{2}", _compact_text(s))]
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
            else:
                amount = ""
                tax = ""
        else:
            amount = ""
            tax = ""
    else:
        amount = ""
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


def _parse_invoice_metadata_from_compact_text(compact: str) -> dict[str, str | None]:
    metadata: dict[str, str | None] = {"优化类型": "invoice_cn"}

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

    buyer_name = _regex_first(compact, [r"购买方信息名称[:：]?(.+?)(?=统一社会信用代码|销售方信息)"])
    seller_name = _regex_first(compact, [r"销售方信息名称[:：]?(.+?)(?=统一社会信用代码|项目名称|货物或应税劳务)"])

    buyer_block = ""
    seller_block = ""
    bi = compact.find("购买方信息")
    si = compact.find("销售方信息")
    if bi != -1:
        end = si if (si != -1 and si > bi) else len(compact)
        buyer_block = compact[bi:end]
    if si != -1:
        end_candidates = []
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

    check_code_raw = _regex_first(compact, [r"校验码[:：]?([0-9OolIZzsB]{20})"])
    check_code = _normalize_ocr_digits(check_code_raw)

    if invoice_kind:
        metadata["发票种类"] = invoice_kind
    if invoice_code:
        metadata["发票代码"] = invoice_code
    if invoice_number:
        metadata["发票号码"] = invoice_number
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

    buyer_bank = _regex_first(
        buyer_block,
        [
            r"开户行及账号[:：]?(.+?)(?=销售方信息|项目名称|货物或应税劳务|$)",
        ],
    )
    seller_bank = _regex_first(
        seller_block,
        [
            r"开户行及账号[:：]?(.+?)(?=项目名称|货物或应税劳务|$)",
        ],
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

    if not invoice_code or not invoice_number or not check_code:
        candidates_raw = re.findall(r"[0-9OolIZzsB]{8,30}", compact)
        candidates = []
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
                issue_date = f"{m.group(1)}年{int(m.group(2)):02d}月{int(m.group(3)):02d}日"
                metadata["开票日期"] = issue_date
                date_digits = yyyymmdd
                break

        if not check_code:
            for d in candidates:
                if len(d) == 20:
                    check_code = d
                    metadata["校验码"] = d
                    break

        if not invoice_code:
            for d in candidates:
                if len(d) in {10, 11, 12}:
                    invoice_code = d
                    metadata["发票代码"] = d
                    break

        if not invoice_number:
            for d in candidates:
                if len(d) == 8 and d != date_digits:
                    invoice_number = d
                    metadata["发票号码"] = d
                    break
        if not invoice_number and date_digits:
            for d in candidates:
                if date_digits in d:
                    rest = d.replace(date_digits, "")
                    if len(rest) == 8:
                        invoice_number = rest
                        metadata["发票号码"] = rest
                        break

    return metadata


def _extract_ofd_text(file_path: str) -> str:
    return _extract_ofd_text_from_items(_extract_ofd_items(file_path))


def _parse_invoice_metadata_from_ofd_items(items: list[tuple[float, float, str]]) -> dict[str, str | None]:
    text = _extract_ofd_text_from_items(items)
    compact = _compact_text(text)
    metadata = _parse_invoice_metadata_from_compact_text(compact)

    lines = _group_items_by_y(items, y_tol=2.2)
    split_x: float | None = None
    for line in lines:
        xs = []
        for x, s in sorted(line, key=lambda t: t[0]):
            c = _compact_text(s)
            if "购买方信息" in c or "销售方信息" in c:
                xs.append(x)
        if len(xs) >= 2:
            xs.sort()
            split_x = (xs[0] + xs[-1]) / 2.0
            break

    buyer: str | None = None
    seller: str | None = None
    buyer_tax_id: str | None = None
    seller_tax_id: str | None = None
    if split_x is not None:
        for line in lines:
            parts = sorted(line, key=lambda t: t[0])
            for idx, (x, s) in enumerate(parts):
                c = _compact_text(s)
                if c not in {"名称：", "名称:"}:
                    continue
                if idx + 1 >= len(parts):
                    continue
                _nx, ns = parts[idx + 1]
                nc = _compact_text(ns)
                if not nc or "项目名称" in nc or "统一社会信用代码" in nc or "纳税人识别号" in nc:
                    continue
                name = str(ns).strip()
                if not name:
                    continue
                if x < split_x and buyer is None:
                    buyer = name
                if x >= split_x and seller is None:
                    seller = name
            if buyer and seller:
                break

        def normalize_tax_id(v: str) -> str:
            v = _compact_text(v).upper()
            v = re.sub(r"[^0-9A-Z]", "", v)
            return v

        def try_pick_tax_id(text: str) -> str | None:
            v = normalize_tax_id(text)
            if re.fullmatch(r"[0-9A-Z]{15,20}", v):
                return v
            return None

        for line in lines:
            parts = sorted(line, key=lambda t: t[0])
            for idx, (x, s) in enumerate(parts):
                side = "buyer" if x < split_x else "seller"
                c = _compact_text(s)
                if "统一社会信用代码" not in c and "纳税人识别号" not in c:
                    continue
                if "项目名称" in c:
                    continue
                tail = ""
                if "：" in str(s) or ":" in str(s):
                    tail = re.split(r"[:：]", str(s), maxsplit=1)[-1]
                picked = try_pick_tax_id(tail) if tail else None
                if picked is None and idx + 1 < len(parts):
                    nx, ns = parts[idx + 1]
                    if (nx < split_x) == (x < split_x):
                        picked = try_pick_tax_id(str(ns))
                if not picked:
                    continue
                if side == "buyer" and buyer_tax_id is None:
                    buyer_tax_id = picked
                if side == "seller" and seller_tax_id is None:
                    seller_tax_id = picked
            if buyer_tax_id and seller_tax_id:
                break

    if buyer:
        metadata["购买方名称"] = buyer
    if seller:
        metadata["销售方名称"] = seller
    if buyer_tax_id:
        metadata["购买方纳税人识别号"] = buyer_tax_id
    if seller_tax_id:
        metadata["销售方纳税人识别号"] = seller_tax_id

    return metadata


def _extract_ofd_items(file_path: str) -> list[tuple[float, float, str]]:
    items: list[tuple[float, float, str]] = []
    with zipfile.ZipFile(file_path) as zf:
        for name in zf.namelist():
            lower = name.lower()
            if not lower.endswith("content.xml"):
                continue
            if "/pages/" not in lower and "/tpls/" not in lower:
                continue
            try:
                xml_bytes = zf.read(name)
            except Exception:
                continue
            try:
                root = etree.fromstring(xml_bytes)
            except Exception:
                continue

            for text_obj in root.iter():
                if _xml_local_name(text_obj.tag) != "TextObject":
                    continue
                boundary = (text_obj.get("Boundary") or "").strip()
                if not boundary:
                    continue
                parts = boundary.split()
                if len(parts) < 2:
                    continue
                try:
                    bx = float(parts[0])
                    by = float(parts[1])
                except Exception:
                    continue

                for code in text_obj.iter():
                    if _xml_local_name(code.tag) != "TextCode":
                        continue
                    s = "".join(code.itertext()).strip()
                    if not s:
                        continue
                    try:
                        cx = float((code.get("X") or "0").strip() or "0")
                        cy = float((code.get("Y") or "0").strip() or "0")
                    except Exception:
                        cx = 0.0
                        cy = 0.0
                    items.append((by + cy, bx + cx, s))

    items.sort(key=lambda t: (t[0], t[1]))
    return items


def _group_items_by_y(
    items: list[tuple[float, float, str]],
    *,
    y_tol: float,
) -> list[list[tuple[float, str]]]:
    lines: list[list[tuple[float, str]]] = []
    current_line: list[tuple[float, str]] = []
    current_y: float | None = None
    for y, x, s in items:
        s = (s or "").strip()
        if not s:
            continue
        if current_y is None:
            current_y = y
            current_line = [(x, s)]
            continue
        if abs(y - current_y) <= y_tol:
            current_line.append((x, s))
            current_y = (current_y + y) / 2.0
            continue
        lines.append(current_line)
        current_y = y
        current_line = [(x, s)]
    if current_line:
        lines.append(current_line)
    return lines


def _extract_ofd_text_from_items(items: list[tuple[float, float, str]]) -> str:
    if not items:
        return ""
    lines = _group_items_by_y(items, y_tol=2.2)
    out_lines: list[str] = []
    for line in lines:
        line_sorted = sorted(line, key=lambda t: t[0])
        out_lines.append("".join([s for _x, s in line_sorted]).strip())
    return "\n".join([l for l in out_lines if l])


def _parse_invoice_rows_from_ofd_items(items: list[tuple[float, float, str]]) -> list[dict[str, str]]:
    header_aliases = {
        "商品名称": {"项目名称", "货物或应税劳务、服务名称", "货物或应税劳务服务名称"},
        "规格型号": {"规格型号"},
        "单位": {"单位", "单 位"},
        "数量": {"数量", "数 量"},
        "单价": {"单价", "单 价"},
        "金额": {"金额"},
        "税率": {"税率/征收率", "税率"},
        "税额": {"税额", "税 额"},
    }

    label_to_key: dict[str, str] = {}
    for key, aliases in header_aliases.items():
        for a in aliases:
            label_to_key[_compact_text(a)] = key

    if not items:
        return []

    lines = _group_items_by_y(items, y_tol=2.2)
    header_idx: int | None = None
    columns: dict[str, float] = {}
    for i, line in enumerate(lines):
        tokens = [s for _x, s in sorted(line, key=lambda t: t[0])]
        if not tokens:
            continue
        compact_line = _compact_text("".join(tokens))
        if "项目名称" not in compact_line and "货物或应税劳务" not in compact_line:
            continue
        if "金额" not in compact_line or "税" not in compact_line:
            continue

        xs = [x for x, _s in sorted(line, key=lambda t: t[0])]
        ss = [s for _x, s in sorted(line, key=lambda t: t[0])]

        for idx in range(len(ss)):
            merged = ""
            for j in range(idx, min(idx + 4, len(ss))):
                merged += _compact_text(ss[j])
                if merged in label_to_key:
                    key = label_to_key[merged]
                    columns.setdefault(key, xs[idx])
                    break

        if len(columns) >= 4 and "商品名称" in columns:
            header_idx = i
            break

    if header_idx is None:
        return []

    ordered = sorted(columns.items(), key=lambda kv: kv[1])
    keys = [k for k, _ in ordered]
    x_starts = [x for _, x in ordered]
    x_bounds = [(x_starts[i] + x_starts[i + 1]) / 2.0 for i in range(len(x_starts) - 1)]

    def assign_col(x0: float) -> str | None:
        if not x_bounds:
            return keys[0]
        idx = bisect_right(x_bounds, x0)
        return keys[idx]

    row_keys = ["商品名称", "规格型号", "单位", "数量", "单价", "金额", "税率", "税额"]
    rows: list[dict[str, str]] = []
    current = {k: "" for k in row_keys}

    def append(k: str, v: str) -> None:
        v = (v or "").strip()
        if not v:
            return
        if current.get(k):
            current[k] = (current[k] + v).strip()
        else:
            current[k] = v

    def flush() -> None:
        nonlocal current
        if any((current.get("商品名称") or "").strip() for _k in ["商品名称"]):
            rows.append({k: (current.get(k) or "").strip() for k in row_keys})
        current = {k: "" for k in row_keys}

    star_chars = ("*", "＊", "∗", "﹡")
    saw_marker = False

    for line in lines[header_idx + 1 :]:
        line_sorted = sorted(line, key=lambda t: t[0])
        if any("合计" in _compact_text(s) for _x, s in line_sorted):
            break

        cols: dict[str, str] = {}
        for x, s in line_sorted:
            col = assign_col(x)
            if not col:
                continue
            cols[col] = (cols.get(col, "") + s).strip()

        if not cols:
            continue

        goods_cell = (cols.get("商品名称") or "").strip()
        if goods_cell.startswith(star_chars):
            saw_marker = True
            if (current.get("商品名称") or "").strip():
                flush()

        for k, v in cols.items():
            if k in current:
                append(k, v)

        if (not saw_marker) and current.get("税额") and (current.get("金额") or current.get("数量")):
            flush()

    if any(v.strip() for v in current.values()):
        flush()

    return [r for r in rows if any((r.get(k) or "").strip() for k in ["商品名称", "金额", "数量"])]


def _parse_invoice_rows_from_pdf_spans(
    *,
    spans: list[tuple[float, float, float, float, str]],
) -> list[dict[str, str]]:
    header_aliases = {
        "商品名称": {"项目名称", "货物或应税劳务、服务名称", "货物或应税劳务服务名称"},
        "规格型号": {"规格型号"},
        "单位": {"单位", "单 位"},
        "数量": {"数量", "数 量"},
        "单价": {"单价", "单 价"},
        "金额": {"金额"},
        "税率": {"税率/征收率", "税率", "税率/征收率"},
        "税额": {"税额", "税 额"},
    }

    label_to_key: dict[str, str] = {}
    for key, aliases in header_aliases.items():
        for a in aliases:
            label_to_key[_compact_text(a)] = key

    columns: dict[str, float] = {}
    header_y0: float | None = None
    for x0, y0, _x1, _y1, s in spans:
        c = _compact_text(s)
        if c in label_to_key:
            key = label_to_key[c]
            columns.setdefault(key, x0)
            if header_y0 is None or y0 < header_y0:
                header_y0 = y0

    if header_y0 is None or len(columns) < 4:
        return []

    ordered = sorted(columns.items(), key=lambda kv: kv[1])
    keys = [k for k, _ in ordered]
    x_starts = [x for _, x in ordered]
    x_bounds = [(x_starts[i] + x_starts[i + 1]) / 2.0 for i in range(len(x_starts) - 1)]

    footer_y0: float | None = None
    for x0, y0, _x1, _y1, s in spans:
        c = _compact_text(s)
        if "价税合计" in c:
            footer_y0 = y0 if footer_y0 is None else min(footer_y0, y0)
            continue
        x_cap = (x_starts[1] + 5) if len(x_starts) >= 2 else (x_starts[0] + 120)
        if (("合计" in c) or (c in {"合", "计"})) and y0 > header_y0 + 10 and x0 <= x_cap:
            footer_y0 = y0 if footer_y0 is None else min(footer_y0, y0)

    def assign_col(x0: float) -> str | None:
        if not x_bounds:
            return keys[0]
        idx = bisect_right(x_bounds, x0)
        return keys[idx]

    def is_header_span(s: str) -> bool:
        return _compact_text(s) in label_to_key

    entries: list[tuple[float, float, str, str]] = []
    for x0, y0, _x1, _y1, s in spans:
        if y0 <= header_y0 + 0.5:
            continue
        if footer_y0 is not None and y0 >= footer_y0 - 2.0:
            continue
        if is_header_span(s):
            continue
        t = (s or "").strip()
        if t in {"¥", "*", "＊", "∗", "﹡"}:
            continue
        col = assign_col(x0)
        if not col:
            continue
        entries.append((float(y0), float(x0), col, t))

    entries.sort(key=lambda t: (t[0], t[1]))

    lines_sorted: list[dict[str, str]] = []
    y_tol = 1.2
    current_y: float | None = None
    line_bucket: dict[str, list[tuple[float, str]]] = {}
    for y0, x0, col, t in entries:
        if current_y is None or abs(y0 - current_y) > y_tol:
            if line_bucket:
                cols = {
                    k: " ".join([s for _x, s in sorted(v, key=lambda it: it[0])]).strip()
                    for k, v in line_bucket.items()
                }
                cols = {k: v for k, v in cols.items() if v}
                if cols:
                    lines_sorted.append(cols)
            current_y = y0
            line_bucket = {}
        else:
            current_y = (current_y + y0) / 2.0
        line_bucket.setdefault(col, []).append((x0, t))
    if line_bucket:
        cols = {k: " ".join([s for _x, s in sorted(v, key=lambda it: it[0])]).strip() for k, v in line_bucket.items()}
        cols = {k: v for k, v in cols.items() if v}
        if cols:
            lines_sorted.append(cols)

    row_keys = ["商品名称", "规格型号", "单位", "数量", "单价", "金额", "税率", "税额"]
    rows: list[dict[str, str]] = []
    row_bucket = {k: "" for k in row_keys}

    def flush() -> None:
        nonlocal row_bucket
        if any((row_bucket.get("商品名称") or "").strip() for _k in ["商品名称"]):
            rows.append({k: (row_bucket.get(k) or "").strip() for k in row_keys})
        row_bucket = {k: "" for k in row_keys}

    def append(k: str, v: str) -> None:
        v = (v or "").strip()
        if not v:
            return
        if row_bucket.get(k):
            if k in {"商品名称", "规格型号"}:
                row_bucket[k] = (row_bucket[k] + v).strip()
            else:
                row_bucket[k] = (row_bucket[k] + " " + v).strip()
        else:
            row_bucket[k] = v

    star_chars = ("*", "＊", "∗", "﹡")
    saw_marker = False

    for cols in lines_sorted:
        goods_cell = (cols.get("商品名称") or "").strip()
        if goods_cell.startswith(star_chars):
            saw_marker = True
            if (row_bucket.get("商品名称") or "").strip():
                flush()

        for k, v in cols.items():
            if k in row_bucket:
                append(k, v)

        if (not saw_marker) and row_bucket.get("税额") and (row_bucket.get("金额") or row_bucket.get("数量")):
            flush()

    if any(v.strip() for v in row_bucket.values()):
        flush()

    return [r for r in rows if any((r.get(k) or "").strip() for k in ["商品名称", "金额", "数量"])]


def _parse_invoice_rows_from_pdf_text(text: str, *, prefer_marked: bool = False) -> list[dict[str, str]]:
    strip_prefix = "\ufeff\u200b\u200c\u200d"
    lines = [((l or "").strip().lstrip(strip_prefix)) for l in (text or "").splitlines()]
    lines = [l for l in lines if l]
    star_chars = ("*", "＊", "∗", "﹡")

    header_aliases = {
        "项目名称",
        "货物或应税劳务、服务名称",
        "货物或应税劳务服务名称",
        "规格型号",
        "单位",
        "数量",
        "单价",
        "金额",
        "税率/征收率",
        "税率",
        "税额",
    }
    headers = {_compact_text(h) for h in header_aliases}

    start = None
    for i, l in enumerate(lines):
        c = _compact_text(l)
        if c in headers or "项目名称" in c or "货物或应税劳务" in c:
            start = i
            break
    if start is None:
        return []

    data_lines: list[str] = []
    seen_header = False
    for l in lines[start:]:
        c = _compact_text(l)
        if c in headers or "项目名称" in c or "货物或应税劳务" in c:
            seen_header = True
            continue
        if not seen_header:
            continue
        if "合计" in _compact_text(l):
            break
        marker_split: list[str] | None = None
        for m in star_chars:
            if m in l and not l.startswith(m) and l.count(m) >= 2:
                idx = l.find(m)
                prefix = l[:idx].strip()
                rest = l[idx:].strip()
                marker_split = []
                if prefix:
                    marker_split.append(prefix)
                if rest:
                    marker_split.append(rest)
                break
        data_lines.extend(marker_split or [l])

    merged_lines: list[str] = []
    i = 0
    while i < len(data_lines):
        t = data_lines[i]
        if i + 1 < len(data_lines) and t == "免" and data_lines[i + 1] == "税":
            merged_lines.append("免税")
            i += 2
            continue
        if i + 2 < len(data_lines) and t == "不" and data_lines[i + 1] == "征" and data_lines[i + 2] == "税":
            merged_lines.append("不征税")
            i += 3
            continue
        c0 = _compact_text(t)
        if len(c0) == 1 and c0 in "0123456789.%％" and i + 1 < len(data_lines):
            j = i
            s = t
            while j + 1 < len(data_lines):
                c1 = _compact_text(data_lines[j + 1])
                if len(c1) == 1 and c1 in "0123456789.%％":
                    s += data_lines[j + 1]
                    j += 1
                    continue
                break
            merged_lines.append(s)
            i = j + 1
            continue
        merged_lines.append(t)
        i += 1

    data_lines = merged_lines

    def is_number(s: str) -> bool:
        s = (s or "").strip().replace("¥", "").replace(",", "")
        return bool(re.fullmatch(r"[0-9]+(?:\.[0-9]{1,4})?", s))

    def is_tax_rate(s: str) -> bool:
        c = _compact_text(s)
        return c in {"免税", "不征税"} or bool(re.fullmatch(r"[0-9]{1,2}(?:\.[0-9]+)?%", c))

    def is_tax_amount(s: str) -> bool:
        c = _compact_text(s)
        return bool(re.fullmatch(r"[*＊\\-—–·•]+", c)) or is_number(c)

    def is_qty(s: str) -> bool:
        return is_number(s)

    def is_unit(s: str) -> bool:
        c = _compact_text(s)
        if not c or len(c) > 3:
            return False
        if is_number(c) or is_tax_rate(c) or bool(re.fullmatch(r"[*＊\\-—–·•]+", c)):
            return False
        return True

    def looks_like_spec(s: str) -> bool:
        c = _compact_text(s)
        if "×" in c or "x" in c.lower():
            return True
        return bool(re.fullmatch(r"[0-9]+(?:\.[0-9]+)?(kg|g|mg|ml|l|L|盒|袋|支|个|片|张)", c))

    row_keys = ["商品名称", "规格型号", "单位", "数量", "单价", "金额", "税率", "税额"]
    rows: list[dict[str, str]] = []

    def explode_tokens(tokens: list[str]) -> list[str]:
        out: list[str] = []
        for t in tokens:
            c = _compact_text(t)
            m = re.fullmatch(r"(免税|不征税)([*＊]+)", c)
            if m:
                out.append(m.group(1))
                out.append(m.group(2))
                continue
            m = re.fullmatch(r"([0-9]{1,2}(?:\.[0-9]+)?%)([*＊]+)", c)
            if m:
                out.append(m.group(1))
                out.append(m.group(2))
                continue
            out.append(t)
        return out

    def build_row(chunk: list[str]) -> dict[str, str] | None:
        chunk = explode_tokens(chunk)
        end = None
        for i in range(len(chunk) - 1, -1, -1):
            if is_tax_amount(chunk[i]):
                end = i
                break
        if end is None:
            return None

        rate_idx = None
        for j in range(end - 1, -1, -1):
            if is_tax_rate(chunk[j]):
                rate_idx = j
                break
        if rate_idx is None:
            return None

        tax_amount = (chunk[end] or "").strip()
        tax_rate = (chunk[rate_idx] or "").strip()

        parts = [chunk[i] for i in range(0, end) if i != rate_idx]
        tail_extras = [s for s in chunk[end + 1 :] if (s or "").strip()]
        amount = parts.pop().strip() if parts and is_number(parts[-1]) else ""
        unit_price = parts.pop().strip() if parts and is_number(parts[-1]) else ""
        qty = parts.pop().strip() if parts and is_qty(parts[-1]) else ""
        unit = parts.pop().strip() if parts and is_unit(parts[-1]) else ""

        parts = [p for p in parts if (p or "").strip()] + tail_extras

        spec_parts: list[str] = []
        while parts and looks_like_spec(parts[-1]):
            spec_parts.insert(0, parts.pop().strip())
        name = "".join([p.strip() for p in parts if p.strip()]).strip()
        spec = " ".join([p for p in spec_parts if p]).strip()

        row = {
            "商品名称": name,
            "规格型号": spec,
            "单位": unit,
            "数量": qty,
            "单价": unit_price,
            "金额": amount,
            "税率": tax_rate,
            "税额": tax_amount,
        }
        if not any((row.get(k) or "").strip() for k in row_keys):
            return None
        if not (row["商品名称"] or row["金额"] or row["数量"]):
            return None
        return row

    chunks: list[list[str]] = []
    current: list[str] = []
    saw_marker = False
    for l in data_lines:
        if l.startswith(star_chars):
            saw_marker = True
            if current:
                chunks.append(current)
            current = [l]
            continue
        current.append(l)
    if current:
        chunks.append(current)

    if saw_marker:
        rows_marked: list[dict[str, str]] = []
        for c in chunks:
            maybe = build_row(c)
            if maybe:
                rows_marked.append(maybe)

        rows_stream: list[dict[str, str]] = []
        stream_buf: list[str] = []
        for l in data_lines:
            stream_buf.append(l)
            if is_tax_amount(l) and len(stream_buf) >= 2 and is_tax_rate(stream_buf[-2]):
                maybe = build_row(stream_buf)
                if maybe:
                    rows_stream.append(maybe)
                    stream_buf = []

        if prefer_marked:
            return rows_marked or rows_stream
        return rows_stream if len(rows_stream) > len(rows_marked) else rows_marked

    buf: list[str] = []
    for l in data_lines:
        buf.append(l)
        if is_tax_amount(l) and len(buf) >= 2 and is_tax_rate(buf[-2]):
            maybe = build_row(buf)
            if maybe:
                rows.append(maybe)
                buf = []

    return rows


compact_text = _compact_text
parse_invoice_metadata_from_compact_text = _parse_invoice_metadata_from_compact_text
parse_invoice_rows_from_pdf_text = _parse_invoice_rows_from_pdf_text
build_yaml_frontmatter = _build_yaml_frontmatter
render_markdown_table = _render_markdown_table


def _parse_pdf_invoice(file_path: str) -> tuple[dict[str, str | None], list[dict[str, str]]]:
    text, spans = _read_pdf_text_and_spans(file_path)
    return _parse_pdf_invoice_from_text_and_spans(text, spans)


_OUTPUT_STEM_TS_RE = re.compile(r"(_\d{8}_\d{6})(.*)$")


def _insert_section_before_timestamp(stem: str, section: str) -> str:
    m = _OUTPUT_STEM_TS_RE.search(stem)
    if not m:
        return f"{stem}_{section}"
    prefix = stem[: m.start()]
    ts = m.group(1)
    suffix = m.group(2)
    return f"{prefix}_{section}{ts}{suffix}"


def convert_invoice_cn_layout_to_md(
    *,
    file_path: str,
    actual_format: str,
    output_dir: str,
    basename_for_output: str,
    original_file_stem: str | None = None,
    cancel_event: threading.Event | None = None,
    progress_callback: Any | None = None,
) -> dict[str, Any]:
    if cancel_event and cancel_event.is_set():
        raise InterruptedError("操作已被取消")

    effective_stem = (original_file_stem or "").strip() or Path(file_path).stem

    headers = ["商品名称", "规格型号", "单位", "数量", "单价", "金额", "税率", "税额"]

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_output_folder = Path(temp_dir) / basename_for_output
        temp_output_folder.mkdir(parents=True, exist_ok=True)
        md_filenames: list[str] = []
        line_count = 0
        page_count = 1

        if actual_format == "ofd":
            metadata, rows = _parse_ofd_invoice(file_path)

            metadata_yaml: dict[str, str | None] = {}
            for k in INVOICE_CN_YAML_SCHEMA:
                v = metadata.get(k)
                metadata_yaml[k] = str(v).strip() if v is not None else ""

            yaml_frontmatter = _build_yaml_frontmatter(
                file_stem=effective_stem, metadata=metadata_yaml, include_empty=True
            )
            table_md = _render_markdown_table(headers=headers, rows=rows)
            md_text = yaml_frontmatter + "## 商品明细\n\n" + table_md

            md_filename = f"{basename_for_output}.md"
            (temp_output_folder / md_filename).write_text(md_text, encoding="utf-8")
            md_filenames.append(md_filename)
            line_count = len(rows)
            page_count = 1

        elif actual_format == "pdf":
            page_count = _get_pdf_page_count(file_path)

            if page_count == 1:
                import fitz

                with fitz.open(file_path) as doc:
                    text = str(doc[0].get_text("text"))

                if _is_scanpage(text):
                    from docwen.converter.layout2md.invoice_cn_ocr import (
                        build_invoice_md_text,
                        parse_invoice_from_image,
                    )

                    png_path = str(temp_output_folder / "__page_1.png")
                    _render_pdf_page_to_png(file_path=file_path, page_index=0, png_path=png_path)
                    try:
                        metadata, rows = parse_invoice_from_image(png_path, cancel_event=cancel_event)
                    finally:
                        Path(png_path).unlink(missing_ok=True)

                    md_text = build_invoice_md_text(
                        file_stem=effective_stem,
                        metadata=metadata,
                        rows=rows,
                        include_empty=True,
                    )
                else:
                    metadata, rows = _parse_pdf_invoice(file_path)
                    metadata_yaml = {}
                    for k in INVOICE_CN_YAML_SCHEMA:
                        v = metadata.get(k)
                        metadata_yaml[k] = str(v).strip() if v is not None else ""

                    yaml_frontmatter = _build_yaml_frontmatter(
                        file_stem=effective_stem,
                        metadata=metadata_yaml,
                        include_empty=True,
                    )
                    table_md = _render_markdown_table(headers=headers, rows=rows)
                    md_text = yaml_frontmatter + "## 商品明细\n\n" + table_md

                md_filename = f"{basename_for_output}.md"
                (temp_output_folder / md_filename).write_text(md_text, encoding="utf-8")
                md_filenames.append(md_filename)
                line_count = len(rows)
            else:
                import fitz

                with fitz.open(file_path) as doc:
                    for page_idx, page in enumerate(doc):
                        if cancel_event and cancel_event.is_set():
                            raise InterruptedError("操作已被取消")

                        text = str(page.get_text("text"))
                        spans: list[tuple[float, float, float, float, str]] = []
                        d = cast(dict[str, Any], page.get_text("dict"))
                        for block in d.get("blocks") or []:
                            if not isinstance(block, dict):
                                continue
                            for line in block.get("lines") or []:
                                if not isinstance(line, dict):
                                    continue
                                for span in line.get("spans") or []:
                                    if not isinstance(span, dict):
                                        continue
                                    s = (span.get("text") or "").strip()
                                    if not s:
                                        continue
                                    x0, y0, x1, y1 = span.get("bbox") or (0.0, 0.0, 0.0, 0.0)
                                    spans.append((float(x0), float(y0), float(x1), float(y1), s))

                        page_num = page_idx + 1
                        section = t("conversion.filenames.page_n", n=page_num)
                        if str(page_num) not in section:
                            section = f"{section}_{page_num}"
                        page_stem_for_yaml = f"{effective_stem}_{section}"
                        page_output_stem = _insert_section_before_timestamp(basename_for_output, section)
                        md_filename = f"{page_output_stem}.md"

                        if _is_scanpage(text):
                            from docwen.converter.layout2md.invoice_cn_ocr import (
                                build_invoice_md_text,
                                parse_invoice_from_image,
                            )

                            png_path = str(temp_output_folder / f"__page_{page_num}.png")
                            _render_pdf_page_to_png(file_path=file_path, page_index=page_idx, png_path=png_path)
                            try:
                                metadata, rows = parse_invoice_from_image(png_path, cancel_event=cancel_event)
                            finally:
                                Path(png_path).unlink(missing_ok=True)
                            md_text = build_invoice_md_text(
                                file_stem=page_stem_for_yaml,
                                metadata=metadata,
                                rows=rows,
                                include_empty=True,
                            )
                            line_count += len(rows)
                        else:
                            metadata, rows = _parse_pdf_invoice_from_text_and_spans(text, spans)
                            metadata_yaml = {}
                            for k in INVOICE_CN_YAML_SCHEMA:
                                v = metadata.get(k)
                                metadata_yaml[k] = str(v).strip() if v is not None else ""

                            yaml_frontmatter = _build_yaml_frontmatter(
                                file_stem=page_stem_for_yaml,
                                metadata=metadata_yaml,
                                include_empty=True,
                            )
                            table_md = _render_markdown_table(headers=headers, rows=rows)
                            md_text = yaml_frontmatter + "## 商品明细\n\n" + table_md
                            line_count += len(rows)

                        (temp_output_folder / md_filename).write_text(md_text, encoding="utf-8")
                        md_filenames.append(md_filename)

        else:
            raise ValueError(f"不支持的发票格式: {actual_format}")

        final_folder, _ = finalize_output(
            str(temp_output_folder),
            str(Path(output_dir) / basename_for_output),
            original_input_file=file_path,
        )
        if not final_folder:
            raise RuntimeError("保存输出文件夹失败")
        if progress_callback:
            try:
                progress_callback(t("conversion.progress.writing_file"))
            except Exception:
                pass
        md_paths = [str(Path(final_folder) / name) for name in md_filenames]
        return {
            "md_path": md_paths[0],
            "md_paths": md_paths,
            "folder_path": final_folder,
            "line_count": line_count,
            "page_count": page_count,
        }


convert_invoice_layout_to_md = convert_invoice_cn_layout_to_md
