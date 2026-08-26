"""OFD invoice parsing through ZIP-container and XML extraction."""

from __future__ import annotations

import re
import xml.etree.ElementTree as etree
import zipfile
from bisect import bisect_right

from docwen_plugin_optimizer_invoice_cn.invoice_cn.ocr_normalize import _compact_text

# ── XML helpers ──────────────────────────────────────────────────────


def _xml_local_name(tag: str) -> str:
    """Strip namespace prefix from an XML tag."""
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def _xml_first_text(root: etree.Element, local_name: str) -> str:
    """Get the first non-empty text content of a named element."""
    for el in root.iter():
        if _xml_local_name(el.tag) != local_name:
            continue
        v = (el.text or "").strip()
        if v:
            return v
    return ""


# ── OFD content item extraction ──────────────────────────────────────


def _extract_ofd_items(
    file_path: str,
) -> list[tuple[float, float, str]]:
    """Extract positioned text items from an OFD file's content.xml pages.

    Args:
        file_path: Path to the OFD file (ZIP container).

    Returns:
        List of ``(y, x, text)`` tuples sorted by position.
    """
    items: list[tuple[float, float, str]] = []
    with zipfile.ZipFile(file_path) as zf:
        for name in zf.namelist():
            lower = name.lower()
            if not lower.endswith("content.xml"):
                continue
            # OFD zip entries may or may not have a leading slash.
            is_page = lower.endswith("content.xml") and ("/pages/" in lower or lower.startswith("pages/"))
            is_tpls = lower.endswith("content.xml") and ("/tpls/" in lower or lower.startswith("tpls/"))
            if not is_page and not is_tpls:
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
    """Group positioned text items into lines by y-coordinate proximity.

    Args:
        items: List of ``(y, x, text)`` tuples sorted by y then x.
        y_tol: Maximum y-difference for items to be considered on the same line.

    Returns:
        List of lines, each line being a list of ``(x, text)`` tuples.
    """
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


def _extract_ofd_text_from_items(
    items: list[tuple[float, float, str]],
) -> str:
    """Build a plain-text representation from OFD positioned text items."""
    if not items:
        return ""
    lines = _group_items_by_y(items, y_tol=2.2)
    out_lines: list[str] = []
    for line in lines:
        line_sorted = sorted(line, key=lambda t: t[0])
        out_lines.append("".join([s for _x, s in line_sorted]).strip())
    return "\n".join(line for line in out_lines if line)


# ── OFD metadata extraction without InvoiceData.xml ──────────────────


def _parse_invoice_metadata_from_ofd_items(
    items: list[tuple[float, float, str]],
) -> dict[str, str | None]:
    """Extract invoice metadata from OFD positioned text items.

    Uses the same compact-text regex approach as PDF parsing,
    augmented with OFD-specific positional heuristics for
    buyer/seller names and tax IDs.
    """
    from docwen_plugin_optimizer_invoice_cn.invoice_cn.metadata import (
        parse_invoice_metadata_from_compact_text,
    )

    text = _extract_ofd_text_from_items(items)
    compact = _compact_text(text)
    metadata = parse_invoice_metadata_from_compact_text(compact)

    lines = _group_items_by_y(items, y_tol=2.2)

    # Detect left/right split from buyer/seller info labels
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


# ── OFD row extraction without InvoiceData.xml ───────────────────────


def _parse_invoice_rows_from_ofd_items(
    items: list[tuple[float, float, str]],
) -> list[dict[str, str]]:
    """Parse invoice detail-line rows from OFD positioned text items.

    Uses header-detection from OFD items to determine column boundaries.
    """
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
        return keys[idx] if idx < len(keys) else None

    row_keys = ["商品名称", "规格型号", "单位", "数量", "单价", "金额", "税率", "税额"]
    rows: list[dict[str, str]] = []
    current = dict.fromkeys(row_keys, "")

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
        current = dict.fromkeys(row_keys, "")

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


# ── Main OFD invoice parser ──────────────────────────────────────────


def parse_ofd_invoice(
    file_path: str,
) -> tuple[dict[str, str | None], list[dict[str, str]]]:
    """Parse an OFD invoice into metadata and detail-line rows.

    Two code paths:
    1. If ``InvoiceData.xml`` is found in the ZIP, parse it directly.
    2. Otherwise, extract positioned text items from OFD content pages
       and use regex+positional heuristics.

    Args:
        file_path: Path to the OFD file (ZIP container).

    Returns:
        A tuple of ``(metadata_dict, rows_list)``.

    Raises:
        ValueError: If the OFD file cannot be parsed.
    """
    # Path 1: Try InvoiceData.xml
    xml_bytes = None
    with zipfile.ZipFile(file_path) as zf:
        for name in zf.namelist():
            if name.lower().endswith("invoicedata.xml"):
                xml_bytes = zf.read(name)
                break

    if not xml_bytes:
        # Path 2: Fallback to OFD content items
        items = _extract_ofd_items(file_path)
        if not items:
            raise ValueError("未找到 InvoiceData.xml 且无法从 OFD 内容提取文本")
        text = _extract_ofd_text_from_items(items)
        metadata = _parse_invoice_metadata_from_ofd_items(items)
        rows = _parse_invoice_rows_from_ofd_items(items)
        if not rows:
            from docwen_plugin_optimizer_invoice_cn.invoice_cn.rows import (
                parse_invoice_rows_from_pdf_text,
            )

            rows = parse_invoice_rows_from_pdf_text(text, prefer_marked=True)
        return metadata, rows

    # Path 1: Parse InvoiceData.xml
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
