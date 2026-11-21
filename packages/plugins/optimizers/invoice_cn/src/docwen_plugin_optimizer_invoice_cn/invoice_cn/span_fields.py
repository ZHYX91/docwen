"""Role-bound invoice metadata within one physical PDF page."""

from __future__ import annotations

import re
from collections.abc import Callable, Sequence
from dataclasses import dataclass
from typing import TYPE_CHECKING
from unicodedata import normalize

from docwen_plugin_optimizer_invoice_cn.invoice_cn.ocr_normalize import (
    NATIVE_AMOUNT_PATTERN,
    _compact_text,
    _normalize_ocr_amount,
    _normalize_positioned_amount,
)
from docwen_plugin_optimizer_invoice_cn.invoice_cn.rows import (
    _assign_column,
    _detect_header_columns,
    _label_text,
    _subtotal_labels,
)

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import ProgressSink

_Item = tuple[float, float, float, str]
_Span = tuple[float, float, float, float, str]
_TAX_LABEL = re.compile(r"(?:统一社会信用代码(?:/纳税人识别号)?|纳税人识别号)[:：]?")
_PERSON_LABELS = ("开票人", "收款人", "复核")
_LOWERCASE = r"(?:\(\s*小写\s*\)|小写)"
_SUBTOTAL_INLINE = re.compile(rf"合\s*计\s*:?\s*({NATIVE_AMOUNT_PATTERN})\s*((?:[+\-−]\s*)?¥\s*(?:[+\-−]\s*)?[0-9].*)")


def report_metadata_issue(sink: ProgressSink | None, *, code: str, location: str, message: str) -> None:
    if sink is not None:
        sink.report_diagnostic("warning", message, code=code, location=location)


@dataclass(frozen=True)
class _PartyRegion:
    role: str
    left: float
    right: float
    top: float
    bottom: float

    def contains(self, x: float, y: float) -> bool:
        return self.left <= x < self.right and self.top <= y < self.bottom


def _party_regions(items: Sequence[_Item]) -> list[_PartyRegion]:
    end = min(
        (
            y
            for y, _x, _end, text in items
            if "项目名称" in _compact_text(text) or "货物或应税劳务" in _compact_text(text)
        ),
        default=float("inf"),
    )
    anchors: dict[str, set[tuple[float, float]]] = {"购买方": set(), "销售方": set()}
    for y, x, _end, text in items:
        compact = _label_text(text)
        roles = [
            role
            for role in anchors
            if re.match(rf"^{role}(?:信息)?(?:[:：]|名称[:：]|统一社会信用代码|纳税人识别号|$)", compact)
        ]
        if len(roles) == 1 and y < end:
            anchors[roles[0]].add((y, x))
    if any(len(values) > 1 for values in anchors.values()):
        return []
    explicit = [(role, *next(iter(values))) for role, values in anchors.items() if values]
    if len(explicit) == 1:
        role, y, _x = explicit[0]
        return [_PartyRegion(role, float("-inf"), float("inf"), y - 2.5, end)]
    if len(explicit) == 2:
        first, second = sorted(explicit, key=lambda item: (item[1], item[2]))
        if abs(first[1] - second[1]) <= 2.5:
            left, right = sorted(explicit, key=lambda item: item[2])
            if left[2] == right[2]:
                return []
            split = right[2] - 2.5
            top = min(left[1], right[1]) - 2.5
            return [
                _PartyRegion(left[0], float("-inf"), split, top, end),
                _PartyRegion(right[0], split, float("inf"), top, end),
            ]
        return [
            _PartyRegion(first[0], float("-inf"), float("inf"), first[1] - 2.5, second[1] - 2.5),
            _PartyRegion(second[0], float("-inf"), float("inf"), second[1] - 2.5, end),
        ]

    # Retain the existing two-column layout evidence, but never treat an
    # unknown split as membership in both parties.
    names = [(y, x) for y, x, _end, text in items if _label_text(text) == "名称:"]
    for index, (y, x) in enumerate(names):
        for other_y, other_x in names[index + 1 :]:
            if abs(other_y - y) <= 2.5 and x != other_x:
                split = max(x, other_x) - 2.5
                return [
                    _PartyRegion("购买方", float("-inf"), split, min(y, other_y) - 2.5, end),
                    _PartyRegion("销售方", split, float("inf"), min(y, other_y) - 2.5, end),
                ]
    tax_x = sorted({x for y, x, _end, text in items if y < end and _TAX_LABEL.search(_compact_text(text))})
    if len(tax_x) >= 3:
        gap, index = max((tax_x[index + 1] - tax_x[index], index) for index in range(len(tax_x) - 1))
        if gap >= 50.0:
            split = (tax_x[index] + tax_x[index + 1]) / 2.0
            return [
                _PartyRegion("购买方", float("-inf"), split, float("-inf"), end),
                _PartyRegion("销售方", split, float("inf"), float("-inf"), end),
            ]
    return []


def _one_value(values: Sequence[str], *, field: str, sink: ProgressSink | None, location: str) -> str:
    unique = set(values)
    if len(unique) == 1:
        return next(iter(unique))
    if len(unique) > 1:
        report_metadata_issue(
            sink,
            code="INVOICE-METADATA-CONFLICT",
            location=f"{location}:{field}",
            message=f"Conflicting {field} values within this page; no positional value was assigned.",
        )
    return ""


def party_metadata(items: Sequence[_Item], *, sink: ProgressSink | None, location: str) -> dict[str, str]:
    regions = _party_regions(items)
    if not regions:
        if any(_TAX_LABEL.search(_compact_text(text)) for _y, _x, _end, text in items):
            report_metadata_issue(
                sink,
                code="INVOICE-PARTY-AMBIGUOUS",
                location=location,
                message="Invoice party labels do not establish distinct buyer/seller regions; unowned positional values were not assigned.",
            )
        return {}

    def role_at(x: float, y: float) -> str:
        roles = [region.role for region in regions if region.contains(x, y)]
        return roles[0] if len(roles) == 1 else ""

    out: dict[str, str] = {}
    for role in ("购买方", "销售方"):
        tax_values: list[str] = []
        names: list[str] = []
        for y, x, _end, text in items:
            if role_at(x, y) != role:
                continue
            compact = _label_text(text)
            tax_label = _TAX_LABEL.search(compact)
            name_label = re.search(r"名\s*称\s*[:：]", text)
            if tax_label is not None:
                tail = compact[tax_label.end() :].upper()
                if re.fullmatch(r"[0-9A-Z]{15,20}", tail):
                    tax_values.append(tail)
                    continue
                candidates: list[tuple[int, str]] = []
                for value_y, value_x, _value_end, value in items:
                    if role_at(value_x, value_y) != role:
                        continue
                    candidate = _label_text(value).upper()
                    if not re.fullmatch(r"[0-9A-Z]{15,20}", candidate):
                        continue
                    if abs(value_y - y) <= 2.5 and value_x > x + 5.0:
                        candidates.append((0, candidate))
                    elif y + 1.0 <= value_y <= y + 25.0 and value_x >= x:
                        candidates.append((1, candidate))
                if candidates:
                    nearest = min(rank for rank, _value in candidates)
                    tax_values.extend(value for rank, value in candidates if rank == nearest)
            elif name_label is not None:
                tail = text[name_label.end() :].strip()
                if tail and not re.search(r"统一社会信用代码|纳税人识别号|名称[:：]", tail):
                    names.append(tail)
                    continue
                if tail:
                    continue
                candidates_text = [
                    (value_x, value.strip())
                    for value_y, value_x, _value_end, value in items
                    if abs(value_y - y) <= 2.5
                    and value_x > x + 5.0
                    and role_at(value_x, value_y) == role
                    and not re.fullmatch(r"[0-9A-Z]{8,30}", _compact_text(value))
                    and not re.search(r"统一社会信用代码|纳税人识别号|名称[:：]|购买方|销售方", value)
                ]
                if candidates_text:
                    names.append(min(candidates_text)[1])
        for field, values in ((role + "纳税人识别号", tax_values), (role + "名称", names)):
            value = _one_value(values, field=field, sink=sink, location=location)
            if value:
                out[field] = value
    return out


def amount_metadata(
    items: Sequence[_Item], spans: list[_Span], *, sink: ProgressSink | None, location: str
) -> dict[str, str]:
    out: dict[str, str] = {}
    grand_labels = [(y, x, end) for y, x, end, text in items if _label_text(text).startswith("价税合计")]
    lowercase_labels = [
        (y, x, end)
        for y, x, end, text in items
        if _label_text(text) in {"小写", "(小写)", "(小写):"}
        and any(abs(y - grand_y) <= 4.2 for grand_y, _x, _end in grand_labels)
    ]
    grand_values = [
        value
        for label_y, label_x, label_end in lowercase_labels or grand_labels
        for y, x, _end, text in items
        if abs(y - label_y) <= 4.2
        and x >= max(label_x + 5, label_end - 6)
        and (value := _normalize_positioned_amount(text))
    ]
    for y, _x, _end, text in items:
        token = normalize("NFKC", text).strip()
        if any(abs(y - grand_y) <= 4.2 for grand_y, _x, _end in grand_labels):
            tail = re.match(
                rf"^(?:价税合计\s*(?:\(大写\).*?{_LOWERCASE}|{_LOWERCASE})?|{_LOWERCASE})\s*:?\s*(.*)$", token
            )
            if tail and (value := _normalize_positioned_amount(tail[1])):
                grand_values.append(value)
    grand = _one_value(grand_values, field="价税合计", sink=sink, location=location)
    if grand:
        out["价税合计"] = grand

    # Only an explicitly labelled total row owns summary amount/tax values.
    # Positive unit prices in an item row cannot stand in for negative totals.
    header = _detect_header_columns(spans)
    left_limit = header[0].get("单位", (float("inf"), 0))[0] if header else float("inf")
    total_labels = _subtotal_labels(spans, left_limit=left_limit)
    if header is not None:
        total_labels = [(y, x) for y, x in total_labels if y > header[1] + 5]
    if grand_labels:
        first_grand_y = min(y for y, _x, _end in grand_labels)
        total_labels = [(y, x) for y, x in total_labels if y < first_grand_y - 2.5]
    inline_totals = []
    for y, x, _end, text in items:
        if x >= left_limit or (header is not None and y <= header[1] + 5):
            continue
        match = _SUBTOTAL_INLINE.fullmatch(normalize("NFKC", text.strip()))
        if match is None or not _normalize_positioned_amount(match[2]):
            continue
        # The whole line matched the native grammar. Monetary normalization
        # can now remove permitted currency spacing.
        inline_totals.append((y, _normalize_ocr_amount(match[1]), _normalize_ocr_amount(match[2])))
    total_y, total_x = max(total_labels, default=(float("-inf"), float("-inf")))
    candidates = [
        (x, end, value)
        for y, x, end, text in items
        if abs(y - total_y) <= 4.2 and x > total_x + 5.0 and (value := _normalize_positioned_amount(text))
    ]
    latest_total_y = max([total_y, *(y for y, _, _ in inline_totals)])
    assigned: dict[str, list[str]] = {"金额": [], "税额": []}
    for y, amount, tax in inline_totals:
        if latest_total_y - y <= 4.2:
            assigned["金额"].append(amount)
            assigned["税额"].append(tax)

    if latest_total_y - total_y > 4.2:
        candidates = []
    if header is not None and {"金额", "税额"}.issubset(header[0]):
        for x, end, value in candidates:
            field = _assign_column(header[0], x, end, value)
            if field in assigned:
                assigned[field].append(value)
    elif len(candidates) == 2 and candidates[0][0] != candidates[1][0]:
        # With no recoverable headers, exactly two values on the explicit
        # total line retain their native left-to-right amount/tax roles.
        ordered_values = sorted(candidates)
        assigned["金额"].append(ordered_values[0][2])
        assigned["税额"].append(ordered_values[1][2])
    elif candidates:
        report_metadata_issue(
            sink,
            code="INVOICE-AMOUNT-AMBIGUOUS",
            location=location,
            message="The total row does not uniquely establish amount/tax columns; no positional monetary fields were guessed.",
        )
    for field, values in assigned.items():
        if value := _one_value(values, field=field, sink=sink, location=location):
            out[field] = value
    return out


def labelled_span_values(
    items: Sequence[_Item],
    patterns: dict[str, tuple[str, Callable[[str], str]]],
    *,
    sink: ProgressSink | None,
    location: str,
) -> dict[str, str]:
    """Read identity and personnel values paired with explicit field labels."""
    labels = (*patterns, *_PERSON_LABELS)
    keyed_items = []
    for y, x, end, text in items:
        if field := next((key for key in labels if _label_text(text).startswith(key)), ""):
            keyed_items.append((field, y, x, end, text))

    out = {}
    for field in labels:
        values = []
        pattern, normalize_value = patterns.get(field, ("", str.strip))
        for key, y, x, end, text in keyed_items:
            if key != field:
                continue
            inline = re.sub(rf"^{field}\s*[:：]?\s*", "", text)
            if inline == text:
                inline = _label_text(text)[len(field) :].lstrip(":")
            if inline.strip():
                candidates = [inline]
            else:
                next_label_x = min(
                    (x2 for _key, y2, x2, _end, _text in keyed_items if abs(y2 - y) <= 2.5 and x2 > x + 5),
                    default=float("inf"),
                )
                candidates = [
                    value
                    for y2, x2, _end, value in items
                    if abs(y2 - y) <= 2.5
                    and max(x + 5, end - 6) <= x2 < next_label_x
                    and not _label_text(value).endswith(":")
                ]
            for candidate in candidates:
                if pattern:
                    candidate = _label_text(candidate)
                    if not re.fullmatch(pattern, candidate):
                        continue
                elif not candidate.strip() or re.fullmatch(r"[:：]+", candidate):
                    continue
                if value := normalize_value(candidate):
                    values.append(value)
        value = _one_value(values, field=field, sink=sink, location=location)
        if value:
            out[field] = value
    return out


def remark_metadata(items: Sequence[_Item]) -> str:
    """Read the remark box, which can have a two-character vertical label."""
    for _y, _x, _end, text in items:
        if match := re.match(r"^备\s*注\s*[:：]\s*(.+)$", text):
            return match[1].strip()

    labels = [(y, x, end) for y, x, end, text in items if _label_text(text) in {"备注", "备注:", "备"}]
    label = next(
        (
            (y, x, end)
            for y, x, end in labels
            if any(
                _label_text(text) in {"备注", "备注:"} and abs(y - ty) < 0.1 and abs(x - tx) < 0.1
                for ty, tx, _end, text in items
            )
            or any(_label_text(text) == "注" and abs(x - tx) <= 2 and y < ty <= y + 30 for ty, tx, _end, text in items)
        ),
        None,
    )
    if label is None:
        return ""
    y, x, _end = label
    grand_y = max(
        (ty for ty, _tx, _end, text in items if ty < y and _label_text(text).startswith("价税合计")), default=y - 5.0
    )
    person_y = min(
        (ty for ty, _tx, _end, text in items if ty > y and _label_text(text).startswith(_PERSON_LABELS)),
        default=y + 25.0,
    )
    values = [
        (ty, tx, text)
        for ty, tx, _end, text in items
        if grand_y + 4.2 < ty < person_y - 2.5
        and tx >= x + 4
        and _label_text(text) not in {"备注", "备注:", "备", "注"}
    ]
    lines: list[list[tuple[float, float, str]]] = []
    for item in sorted(set(values)):
        if not lines or item[0] - lines[-1][0][0] > 2.0:
            lines.append([])
        lines[-1].append(item)
    return "\n".join("".join(text for _y, _x, text in sorted(line, key=lambda part: part[1])) for line in lines)
