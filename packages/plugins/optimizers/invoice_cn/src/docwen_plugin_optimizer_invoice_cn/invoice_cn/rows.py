"""Invoice detail-line row parsing from positioned spans and text."""

from __future__ import annotations

import re
from bisect import bisect_right
from itertools import pairwise
from typing import TYPE_CHECKING
from unicodedata import normalize

from docwen_plugin_optimizer_invoice_cn.invoice_cn.ocr_normalize import _compact_text, _normalize_positioned_amount
from docwen_plugin_optimizer_invoice_cn.invoice_cn.yaml_schema import TABLE_HEADERS

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import ProgressSink

_HEADER_ALIASES: dict[str, set[str]] = {
    "商品名称": {"商品名称", "项目名称", "货物或应税劳务、服务名称", "货物或应税劳务服务名称"},
    "规格型号": {"规格型号"},
    "单位": {"单位"},
    "数量": {"数量"},
    "单价": {"单价"},
    "金额": {"金额"},
    "税率": {"税率/征收率", "税率征收率", "税率"},
    "税额": {"税额"},
}
_ROW_KEYS = TABLE_HEADERS
_Span = tuple[float, float, float, float, str]
_ColumnGeometry = dict[str, tuple[float, float]]
_NUMERIC_COLUMNS = {"数量", "单价", "金额", "税额"}
_TAX_RATE = re.compile(r"(?:免税|不征税|[0-9]{1,2}(?:\.[0-9]+)?%)")
_STARS = ("*", "＊", "∗", "﹡")
_EXEMPT_TAX = re.compile(r"[*＊∗﹡\-—–]+")


def _label_text(value: str) -> str:
    return normalize("NFKC", _compact_text(value))


def _detect_header_columns(spans: list[_Span]) -> tuple[_ColumnGeometry, float] | None:
    """Recover both edges of each label, including headers split across spans."""
    groups: list[list[_Span]] = []
    for span in sorted(spans, key=lambda item: (item[1], item[0])):
        if not groups or abs(span[1] - groups[-1][0][1]) > 1.5:
            groups.append([])
        groups[-1].append(span)
    best: tuple[_ColumnGeometry, float] | None = None
    for group in groups:
        characters: list[tuple[float, float, str]] = []
        for x0, _y0, x1, _y1, raw in group:
            advance = max(x1 - x0, 0.0) / max(len(raw), 1)
            for index, char in enumerate(raw):
                for compact in _label_text(char):
                    characters.append((x0 + advance * index, x0 + advance * (index + 1), compact))

        characters.sort(key=lambda item: item[0])
        line = "".join(char for _x0, _x1, char in characters)
        columns = {}
        for key, aliases in _HEADER_ALIASES.items():
            for alias in sorted(aliases, key=len, reverse=True):
                start = line.find(alias)
                if start >= 0:
                    columns[key] = (characters[start][0], characters[start + len(alias) - 1][1])
                    break
        if "商品名称" in columns and len(columns) >= 4 and (best is None or len(columns) > len(best[0])):
            best = (columns, group[0][1])
    return best


def _assign_column(columns: _ColumnGeometry, x0: float, x1: float, text: str) -> str:
    compact = _label_text(text)
    if "税率" in columns and _TAX_RATE.fullmatch(compact) and abs(x0 - columns["税率"][0]) < 30:
        return "税率"
    if "税额" in columns and _EXEMPT_TAX.fullmatch(compact) and x0 >= columns["税额"][0] - 15:
        return "税额"
    ordered = sorted(columns, key=lambda key: columns[key][0])
    starts = [columns[key][0] for key in ordered]
    bounds = [(left + right) / 2 for left, right in pairwise(starts)]
    column = ordered[bisect_right(bounds, x0)]
    if (column in _NUMERIC_COLUMNS or column == "税率") and _normalize_positioned_amount(text):
        candidates = [key for key in ordered if key in _NUMERIC_COLUMNS]
        # Text may be left- or right-aligned. In particular, a long unit
        # price extends across the midpoint between price and quantity.
        return min(
            candidates,
            key=lambda key: (min(abs(x0 - columns[key][0]), abs(x1 - columns[key][1])), abs(x0 - columns[key][0])),
        )
    return column


def _subtotal_labels(spans: list[_Span], *, left_limit: float = float("inf")) -> list[tuple[float, float]]:
    direct = [
        (y, x)
        for x, y, _end, _y1, text in spans
        if x < left_limit and re.match(r"^合计:?(?:$|[+\-−]?[¥￥])", _label_text(text))
    ]
    split = [
        (y, x)
        for x, y, _end, _y1, first in spans
        if x < left_limit
        and _label_text(first) == "合"
        and any(
            abs(y - other_y) <= 2.0 and x < other_x < x + 90 and _label_text(last) in {"计", "计:"}
            for other_x, other_y, _end, _y1, last in spans
        )
    ]
    return direct + split


def parse_invoice_rows_from_pdf_spans(
    *,
    spans: list[_Span],
    diagnostic_sink: ProgressSink | None = None,
    location: str = "",
) -> list[dict[str, str]]:
    """Parse one positioned page, assigning each cell to its closest detail row."""
    header = _detect_header_columns(spans)
    if header is None:
        return []
    columns, header_y = header
    left_limit = columns.get("单位", (columns["商品名称"][0] + 120.0, 0.0))[0]
    subtotal_levels = [y for y, _x in _subtotal_labels(spans, left_limit=left_limit) if y > header_y + 5]
    footer_levels = [
        y
        for _x, y, _end, _y1, text in spans
        if y > header_y and re.match(r"^(?:价税合计|备注|开票人|收款人|复核)", _label_text(text))
    ]
    end = min(subtotal_levels + footer_levels, default=float("inf"))
    if subtotal_levels and end == min(subtotal_levels):
        end -= 4.2

    entries: list[tuple[float, float, float, str, str]] = []
    for x0, y, x1, _y1, raw in set(spans):
        text = raw.strip()
        if not (header_y + 1.5 < y < end and text) or text in {"¥", "￥"}:
            continue
        column = _assign_column(columns, x0, x1, text)
        if _label_text(text) in _HEADER_ALIASES.get(column, ()):
            continue
        entries.append((y, x0, x1, column, text))
    entries.sort()

    anchors: list[float] = []
    for y, _x, _end, column, text in entries:
        if column == "商品名称" and text.startswith(_STARS) and (not anchors or y - anchors[-1] > 2.5):
            anchors.append(y)
    for y, _x, _end, column, text in entries:
        if column not in {"金额", "税额", "税率"}:
            continue
        if not (
            _normalize_positioned_amount(text) or _TAX_RATE.fullmatch(_label_text(text)) or _EXEMPT_TAX.fullmatch(text)
        ):
            continue
        if not any(abs(anchor - y) <= 2.5 for anchor in anchors):
            anchors.append(y)
    anchors.sort()
    if not anchors:
        return []

    buckets: list[dict[str, list[tuple[float, float, float, str]]]] = [{} for _y in anchors]
    for y, x0, x1, column, text in entries:
        index = min(range(len(anchors)), key=lambda index: abs(anchors[index] - y))
        if column == "商品名称" and abs(anchors[index] - y) > 2.5:
            # Wrapped goods descriptions follow their own line's first span.
            index = max(bisect_right(anchors, y) - 1, 0)
        elif column in _NUMERIC_COLUMNS | {"税率", "单位"} and abs(anchors[index] - y) > 2.5:
            continue
        buckets[index].setdefault(column, []).append((y, x0, x1, text))

    rows = []
    for row_number, bucket in enumerate(buckets, start=1):
        row = dict.fromkeys(_ROW_KEYS, "")
        for field, parts in bucket.items():
            values = [text for _y, _x, _end, text in parts]
            if field in _NUMERIC_COLUMNS and len(values) > 1:
                if diagnostic_sink is not None:
                    diagnostic_sink.report_diagnostic(
                        "warning",
                        f"Multiple numeric spans claim {field} on this detail row.",
                        code="INVOICE-ROW-AMBIGUOUS",
                        location=f"{location}:row-{row_number}:{field}",
                    )
                continue
            row[field] = "".join(values)
        if any(row.values()):
            rows.append(row)
    return rows


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
