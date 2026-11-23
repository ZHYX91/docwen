"""Merge stored worksheet records along either axis without expanding blank rectangles."""

from __future__ import annotations

from collections.abc import Callable
from typing import Any

from docwen_plugin_spreadsheet.table_merger.semantics import MergeSummary, value_identity
from docwen_plugin_spreadsheet.table_merger.text_limits import check_cell_text


def write_scalar(cell: Any, value: Any, source_type: str | None = None) -> None:
    """Write saved values without reinterpreting literal text as formulas or errors."""
    if isinstance(value, str):
        check_cell_text(value, sheet=cell.parent.title, coordinate=cell.coordinate)
    cell.value = value
    if isinstance(value, str):
        cell.data_type = "e" if source_type == "e" else "s"


def check_extent(rows: int, columns: int) -> None:
    if rows > 1_048_576 or columns > 16_384:
        raise ValueError("Merged worksheet exceeds Excel's row or column limit")


def insert_sparse(ws: Any, *, by_row: bool, start: int, amount: int, cancel_check: Callable[[], None]) -> None:
    """Move stored cells, including styled blanks, in descending coordinate order."""
    # openpyxl insert_rows/cols materializes the entire affected rectangle first.
    for index, (row, col) in enumerate(sorted(ws._cells, reverse=True)):
        if index % 1000 == 0:
            cancel_check()
        if (row if by_row else col) >= start:
            cell = ws._cells[row, col]
            ws._move_cell(row, col, amount if by_row else 0, 0 if by_row else amount)
            if cell.hyperlink:
                cell.hyperlink.ref = cell.coordinate
    if by_row:
        ws._current_row = ws.max_row


def _stored_records(ws: Any, by_row: bool, offset: int, cancel_check: Callable[[], None]) -> dict[int, dict[int, Any]]:
    records: dict[int, dict[int, Any]] = {}
    for index, ((row, col), cell) in enumerate(ws._cells.items()):
        if index % 1000 == 0:
            cancel_check()
        if cell.value is not None:
            record, field = (row, col) if by_row else (col, row)
            records.setdefault(record, {})[field + offset] = cell
    return records


def _identities(cells: dict[int, Any]) -> dict[int, Any]:
    return {
        field: identity
        for field, cell in cells.items()
        if (identity := value_identity(cell.value, source_type=cell.data_type)) is not None
    }


def _similarity(base: dict[int, Any], incoming: dict[int, Any], width: int, offset: int) -> float:
    """Score the same aligned fields as a dense comparison, counting implicit blank matches arithmetically."""
    outside = sum(not offset < field <= offset + width for field in base)
    differences = sum(value != base.get(field) for field, value in incoming.items())
    differences += sum(field not in incoming and offset < field <= offset + width for field in base)
    return (width - differences) / (width + outside)


def _best_position(
    records: dict[int, dict[int, Any]],
    incoming: dict[int, Any],
    extent: int,
    width: int,
    offset: int,
    cancel_check: Callable[[], None],
) -> tuple[str, int]:
    """Keep the first covering record, or insert after the first most similar record."""
    best, score = 1, 0.0
    next_record = 1
    for sequence, index in enumerate(sorted(records)):
        if sequence % 100 == 0:
            cancel_check()
        if next_record < index:
            # Every empty record has the same score. Only the first can win a tie.
            blank_score = _similarity({}, incoming, width, offset)
            if blank_score > score:
                best, score = next_record, blank_score
        base = records[index]
        if incoming.items() <= base.items():
            return "skipped", index
        if base.items() < incoming.items():
            return "replaced", index
        similarity = _similarity(base, incoming, width, offset)
        if similarity > score:
            best, score = index, similarity
        next_record = index + 1
    if next_record <= extent and _similarity({}, incoming, width, offset) > score:
        best = next_record
    return "added", best + 1


def _skip_empty_records(summary: MergeSummary, label: str, start: int, stop: int) -> None:
    if stop <= start:
        return
    summary.counts["skipped"] += stop - start
    samples = summary.locations.setdefault("skipped", [])
    samples.extend(f"{label} {index}" for index in range(start, min(stop, start + 20 - len(samples))))


def merge_records(
    base: Any, collect: Any, *, by_row: bool, offset: int, summary: MergeSummary, cancel_check: Callable[[], None]
) -> None:
    """Align rows or columns using sparse field identities, preserving source types on writes."""
    source = _stored_records(collect, by_row, offset, cancel_check)
    contents = _stored_records(base, by_row, 0, cancel_check)
    identities = {index: values for index, fields in contents.items() if (values := _identities(fields))}
    width, collect_extent = (collect.max_column, collect.max_row) if by_row else (collect.max_row, collect.max_column)
    base_extent = base.max_row if by_row else base.max_column
    label = "row" if by_row else "column"
    next_record = 1
    for index in sorted(source):
        cancel_check()
        incoming = _identities(source[index])
        if not incoming:
            continue
        _skip_empty_records(summary, label, next_record, index)
        next_record = index + 1
        decision, target = _best_position(identities, incoming, base_extent, width, offset, cancel_check)
        summary.record(decision, f"{label} {index}")
        if decision == "skipped":
            continue
        if decision == "added":
            base_extent += 1
            check_extent(base_extent if by_row else width + offset, width + offset if by_row else base_extent)
            insert_sparse(base, by_row=by_row, start=target, amount=1, cancel_check=cancel_check)
            for records in (identities, contents):
                for key in sorted((key for key in records if key >= target), reverse=True):
                    records[key + 1] = records.pop(key)
        # A replacement may clear stored whitespace where the source has an implicit blank.
        fields = contents.setdefault(target, {})
        for field in fields.keys() - source[index].keys():
            if offset < field <= offset + width:
                write_scalar(fields.pop(field), None)
        for field, cell in source[index].items():
            row, col = (target, field) if by_row else (field, target)
            check_extent(row, col)
            dest = base.cell(row, col)
            write_scalar(dest, cell.value, cell.data_type)
            fields[field] = dest
        identities[target] = incoming
    _skip_empty_records(summary, label, next_record, collect_extent + 1)
