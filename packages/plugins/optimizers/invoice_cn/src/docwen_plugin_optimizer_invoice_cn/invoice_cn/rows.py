"""Invoice detail-line row parsing from PDF spans and text."""

from __future__ import annotations

import re
from bisect import bisect_right

from docwen_plugin_optimizer_invoice_cn.invoice_cn.ocr_normalize import _compact_text

# ── Header aliases for column detection ──────────────────────────────

_HEADER_ALIASES: dict[str, set[str]] = {
    "商品名称": {"项目名称", "货物或应税劳务、服务名称", "货物或应税劳务服务名称"},
    "规格型号": {"规格型号"},
    "单位": {"单位", "单 位"},
    "数量": {"数量", "数 量"},
    "单价": {"单价", "单 价"},
    "金额": {"金额"},
    "税率": {"税率/征收率", "税率征收率", "税率"},
    "税额": {"税额", "税 额"},
}

_ROW_KEYS = ["商品名称", "规格型号", "单位", "数量", "单价", "金额", "税率", "税额"]


def _build_label_map() -> dict[str, str]:
    """Build a label→column-key lookup from header aliases."""
    label_to_key: dict[str, str] = {}
    for key, aliases in _HEADER_ALIASES.items():
        for a in aliases:
            label_to_key[_compact_text(a)] = key
    return label_to_key


def _detect_header_columns(
    spans: list[tuple[float, float, float, float, str]],
) -> tuple[dict[str, float], float, float] | None:
    """Detect one invoice header line, including split or combined spans.

    Some electronic invoices expose adjacent headers as one PDF text span
    (for example ``金额税率/征收率``), while others expose every Chinese
    character as a separate span.  Reconstructing character positions lets the
    parser recover both shapes without deriving columns from invoice values.
    """
    line_groups: list[list[tuple[float, float, float, float, str]]] = []
    line_y: list[float] = []
    for span in sorted(spans, key=lambda item: (float(item[1]), float(item[0]))):
        y0 = float(span[1])
        if not line_groups or abs(y0 - line_y[-1]) > 1.2:
            line_groups.append([span])
            line_y.append(y0)
            continue
        line_groups[-1].append(span)
        line_y[-1] = (line_y[-1] + y0) / 2.0

    aliases_by_key = {
        key: sorted({_compact_text(alias) for alias in aliases}, key=len, reverse=True)
        for key, aliases in _HEADER_ALIASES.items()
    }
    best: tuple[dict[str, float], float, float] | None = None
    best_score = 0

    for group in line_groups:
        characters: list[tuple[float, str]] = []
        for x0, _y0, x1, _y1, text in group:
            raw = text or ""
            if not raw:
                continue
            width = max(float(x1) - float(x0), 0.0)
            for index, char in enumerate(raw):
                compact_char = _compact_text(char)
                if not compact_char:
                    continue
                characters.append((float(x0) + width * index / len(raw), compact_char))

        characters.sort(key=lambda item: item[0])
        compact_line = "".join(char for _x, char in characters)
        if not compact_line:
            continue

        columns: dict[str, float] = {}
        for key, aliases in aliases_by_key.items():
            for alias in aliases:
                start = compact_line.find(alias)
                if start >= 0:
                    columns[key] = characters[start][0]
                    break

        score = len(columns)
        group_y0 = min(float(span[1]) for span in group)
        group_y1 = max(float(span[1]) for span in group)
        if score > best_score or (score == best_score and best is not None and group_y0 < best[1]):
            best = (columns, group_y0, group_y1)
            best_score = score

    if best is None or best_score < 4:
        return None
    return best


# ── Span-based row parsing ───────────────────────────────────────────


def parse_invoice_rows_from_pdf_spans(
    *,
    spans: list[tuple[float, float, float, float, str]],
) -> list[dict[str, str]]:
    """Parse invoice detail-line rows from PDF span positions.

    Uses header-detection from spans to determine column boundaries,
    then assigns each data span to a column based on its x-coordinate.

    Args:
        spans: List of (x0, y0, x1, y1, text) tuples from PyMuPDF.

    Returns:
        List of row dicts with keys matching ``_ROW_KEYS``.
    """
    label_to_key = _build_label_map()
    detected_header = _detect_header_columns(spans)
    if detected_header is None:
        return []
    columns, header_y0, header_y1 = detected_header

    ordered = sorted(columns.items(), key=lambda kv: kv[1])
    keys = [k for k, _ in ordered]
    x_starts = [x for _, x in ordered]
    x_bounds = [(x_starts[i] + x_starts[i + 1]) / 2.0 for i in range(len(x_starts) - 1)]

    # Detect footer row (价税合计 / 合计)
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
        return keys[idx] if idx < len(keys) else None

    def is_header_span(s: str) -> bool:
        return _compact_text(s) in label_to_key

    star_chars = ("*", "＊", "∗", "﹡")

    entries: list[tuple[float, float, str, str]] = []
    for x0, y0, _x1, _y1, s in spans:
        if y0 <= header_y1 + 0.5:
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

    # Group entries into lines by y-coordinate
    y_tol = 1.2
    lines_sorted: list[dict[str, str]] = []
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

    # Merge lines into rows (multi-line cells)
    rows: list[dict[str, str]] = []
    row_bucket = dict.fromkeys(_ROW_KEYS, "")

    def flush() -> None:
        nonlocal row_bucket
        if any((row_bucket.get("商品名称") or "").strip() for _k in ["商品名称"]):
            rows.append({k: (row_bucket.get(k) or "").strip() for k in _ROW_KEYS})
        row_bucket = dict.fromkeys(_ROW_KEYS, "")

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


# ── Text-based row parsing ───────────────────────────────────────────


def parse_invoice_rows_from_pdf_text(text: str, *, prefer_marked: bool = False) -> list[dict[str, str]]:
    """Parse invoice detail-line rows from raw text (line-based).

    Used as a fallback when span-based parsing fails.

    Args:
        text: Raw text from the PDF.
        prefer_marked: If True, prefer rows identified by leading asterisk markers.

    Returns:
        List of row dicts with keys matching ``_ROW_KEYS``.
    """
    strip_prefix = "﻿​‌‍"
    lines = [((line or "").strip().lstrip(strip_prefix)) for line in (text or "").splitlines()]
    lines = [line for line in lines if line]
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

    # Find the header row
    start = None
    for i, line in enumerate(lines):
        c = _compact_text(line)
        if c in headers or "项目名称" in c or "货物或应税劳务" in c:
            start = i
            break
    if start is None:
        return []

    # Collect data lines after the header
    data_lines: list[str] = []
    seen_header = False
    for line in lines[start:]:
        c = _compact_text(line)
        if c in headers or "项目名称" in c or "货物或应税劳务" in c:
            seen_header = True
            continue
        if not seen_header:
            continue
        if "合计" in _compact_text(line):
            break
        marker_split: list[str] | None = None
        for m in star_chars:
            if m in line and not line.startswith(m) and line.count(m) >= 2:
                idx = line.find(m)
                prefix = line[:idx].strip()
                rest = line[idx:].strip()
                marker_split = []
                if prefix:
                    marker_split.append(prefix)
                if rest:
                    marker_split.append(rest)
                break
        data_lines.extend(marker_split or [line])

    # Merge split number-lines (single-digit lines are likely OCR artifacts)
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

    # ── Token classifiers ────────────────────────────────────────────

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
        return not (is_number(c) or is_tax_rate(c) or bool(re.fullmatch(r"[*＊\\-—–·•]+", c)))

    def looks_like_spec(s: str) -> bool:
        c = _compact_text(s)
        if "×" in c or "x" in c.lower():
            return True
        return bool(re.fullmatch(r"[0-9]+(?:\.[0-9]+)?(kg|g|mg|ml|l|L|盒|袋|支|个|片|张)", c))

    # ── Row builder ──────────────────────────────────────────────────

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

        parts = [chunk[i] for i in range(end) if i != rate_idx]
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
        if not any((row.get(k) or "").strip() for k in _ROW_KEYS):
            return None
        if not (row["商品名称"] or row["金额"] or row["数量"]):
            return None
        return row

    # ── Chunk and build ──────────────────────────────────────────────

    chunks: list[list[str]] = []
    current: list[str] = []
    saw_marker = False
    for line in data_lines:
        if line.startswith(star_chars):
            saw_marker = True
            if current:
                chunks.append(current)
            current = [line]
            continue
        current.append(line)
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
        for line in data_lines:
            stream_buf.append(line)
            if is_tax_amount(line) and len(stream_buf) >= 2 and is_tax_rate(stream_buf[-2]):
                maybe = build_row(stream_buf)
                if maybe:
                    rows_stream.append(maybe)
                    stream_buf = []

        if prefer_marked:
            return rows_marked or rows_stream
        return rows_stream if len(rows_stream) > len(rows_marked) else rows_marked

    rows: list[dict[str, str]] = []
    buf: list[str] = []
    for line in data_lines:
        buf.append(line)
        if is_tax_amount(line) and len(buf) >= 2 and is_tax_rate(buf[-2]):
            maybe = build_row(buf)
            if maybe:
                rows.append(maybe)
                buf = []

    return rows
