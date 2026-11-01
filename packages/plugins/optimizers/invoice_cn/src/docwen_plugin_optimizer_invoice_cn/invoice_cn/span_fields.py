"""Role-bound invoice metadata within one physical PDF page."""

from __future__ import annotations

import re
from bisect import bisect_right
from collections.abc import Sequence
from dataclasses import dataclass
from typing import TYPE_CHECKING

from docwen_plugin_optimizer_invoice_cn.invoice_cn.ocr_normalize import _compact_text, _normalize_ocr_amount
from docwen_plugin_optimizer_invoice_cn.invoice_cn.rows import _detect_header_columns

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import ProgressSink

_Item = tuple[float, float, float, str]
_Span = tuple[float, float, float, float, str]
_TAX_LABEL = re.compile(r"(?:统一社会信用代码(?:/纳税人识别号)?|纳税人识别号)[:：]?")


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
        compact = _compact_text(text)
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
            split = (left[2] + right[2]) / 2.0
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
    names = [(y, x) for y, x, _end, text in items if _compact_text(text) in {"名称：", "名称:"}]
    for index, (y, x) in enumerate(names):
        for other_y, other_x in names[index + 1 :]:
            if abs(other_y - y) <= 2.5 and x != other_x:
                split = (x + other_x) / 2.0
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
            compact = _compact_text(text)
            tax_label = _TAX_LABEL.search(compact)
            name_label = re.search(r"名称[:：]", text)
            if tax_label is not None:
                tail = compact[tax_label.end() :].upper()
                if re.fullmatch(r"[0-9A-Z]{15,20}", tail):
                    tax_values.append(tail)
                    continue
                candidates: list[tuple[int, str]] = []
                for value_y, value_x, _value_end, value in items:
                    if role_at(value_x, value_y) != role:
                        continue
                    candidate = _compact_text(value).upper()
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
    grand_labels = [(y, x, end) for y, x, end, text in items if "价税合计" in _compact_text(text)]
    grand_values = [
        value
        for label_y, _label_x, label_end in grand_labels
        for y, x, _end, text in items
        if abs(y - label_y) <= 2.5 and x >= label_end - 2.5 and (value := _normalize_ocr_amount(_compact_text(text)))
    ]
    grand = _one_value(grand_values, field="价税合计", sink=sink, location=location)
    if grand:
        out["价税合计"] = grand

    # Only an explicitly labelled total row owns summary amount/tax values.
    # Positive unit prices in an item row cannot stand in for negative totals.
    total_labels = [(y, x) for y, x, _end, text in items if re.fullmatch(r"合计[:：]?", _compact_text(text))]
    if grand_labels:
        first_grand_y = min(y for y, _x, _end in grand_labels)
        total_labels = [(y, x) for y, x in total_labels if y < first_grand_y - 2.5]
    if not total_labels:
        return out
    total_y, total_x = max(total_labels)
    candidates = [
        (x, value)
        for y, x, _end, text in items
        if abs(y - total_y) <= 2.5 and x > total_x + 5.0 and (value := _normalize_ocr_amount(_compact_text(text)))
    ]
    header = _detect_header_columns(spans)
    if header is not None and {"金额", "税额"}.issubset(header[0]):
        ordered = sorted(header[0].items(), key=lambda item: item[1])
        bounds = [(ordered[index][1] + ordered[index + 1][1]) / 2.0 for index in range(len(ordered) - 1)]
        assigned: dict[str, list[str]] = {"金额": [], "税额": []}
        for x, value in candidates:
            field = ordered[bisect_right(bounds, x)][0]
            if field in assigned:
                assigned[field].append(value)
        for field, values in assigned.items():
            value = _one_value(values, field=field, sink=sink, location=location)
            if value:
                out[field] = value
    elif len(candidates) == 2 and candidates[0][0] != candidates[1][0]:
        # With no recoverable headers, exactly two values on the explicit
        # total line retain their native left-to-right amount/tax roles.
        ordered_values = sorted(candidates)
        out["金额"], out["税额"] = ordered_values[0][1], ordered_values[1][1]
    elif candidates:
        report_metadata_issue(
            sink,
            code="INVOICE-AMOUNT-AMBIGUOUS",
            location=location,
            message="The total row does not uniquely establish amount/tax columns; no positional monetary fields were guessed.",
        )
    return out
