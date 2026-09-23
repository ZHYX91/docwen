"""Offline table localisation and structure recognition sharing admitted OCR.

Only OCR regions consumed by a valid table are replaced. Remaining regions stay
in reading order, including prose and unsuccessful table crops.
"""

from __future__ import annotations

import hashlib
import html
import os
import sys
from collections.abc import Callable, Mapping
from dataclasses import dataclass, replace
from enum import StrEnum
from functools import lru_cache
from html.parser import HTMLParser
from pathlib import Path
from typing import cast

from docwen_core.errors import CancellationRequested
from docwen_core.export_semantics import TableMergeRegion, build_table_semantic_grid, render_table_semantic_grid
from docwen_core.text.ocr import OcrOutcome, OcrStatus

RAPIDTABLE_MODEL_FILENAME = "slanet-plus.onnx"
RAPIDTABLE_MODEL_SHA256 = "d57a942af6a2f57d6a4a0372573c696a2379bf5857c45e2ac69993f3b334514b"
RAPIDTABLE_MODEL_URL = "https://www.modelscope.cn/models/RapidAI/RapidTable/resolve/v2.0.0/slanet-plus.onnx"
TABLE_LAYOUT_MODEL_FILENAME = "layout_table.onnx"
TABLE_LAYOUT_MODEL_SHA256 = "5b07ba6df1d1889bed2877c9d7501235c6fb6e2212aca8f2f56f4b1b8d0e37b5"
TABLE_LAYOUT_MODEL_URL = (
    "https://www.modelscope.cn/models/RapidAI/RapidLayout/resolve/v1.0.0/onnx/pp_layout/layout_table.onnx"
)


class TableRecognitionStatus(StrEnum):
    SUCCESS = "success"
    NOT_TABLE = "not_table"
    DISABLED = "disabled"
    UNAVAILABLE = "unavailable"
    FAILED = "failed"


@dataclass(frozen=True, slots=True)
class TableRecognitionOutcome:
    status: TableRecognitionStatus
    markdown: str = ""
    message: str = ""
    table_count: int = 0

    @property
    def diagnostic_code(self) -> str:
        return "OCR-TABLE-UNAVAILABLE" if self.status is TableRecognitionStatus.UNAVAILABLE else "OCR-TABLE-FALLBACK"

    @property
    def fallback_required(self) -> bool:
        return self.status in {TableRecognitionStatus.UNAVAILABLE, TableRecognitionStatus.FAILED} or bool(self.message)

    @property
    def recognized_table(self) -> bool:
        return self.status is TableRecognitionStatus.SUCCESS and bool(self.markdown)


class _TableHTMLParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.rows: list[list[tuple[str, int, int]]] = []
        self._row: list[tuple[str, int, int]] | None = None
        self._cell_parts: list[str] | None = None
        self._rowspan = 1
        self._colspan = 1

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag == "tr":
            self._row = []
        elif tag in {"td", "th"} and self._row is not None:
            values = dict(attrs)
            self._rowspan = _positive_span(values.get("rowspan"))
            self._colspan = _positive_span(values.get("colspan"))
            self._cell_parts = []
        elif tag == "br" and self._cell_parts is not None:
            self._cell_parts.append("\n")

    def handle_data(self, data: str) -> None:
        if self._cell_parts is not None:
            self._cell_parts.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag in {"td", "th"} and self._cell_parts is not None and self._row is not None:
            self._row.append(("".join(self._cell_parts).strip(), self._rowspan, self._colspan))
            self._cell_parts = None
        elif tag == "tr" and self._row is not None:
            if self._row:
                self.rows.append(self._row)
            self._row = None


def _positive_span(value: str | None) -> int:
    try:
        return max(1, min(int(value or "1"), 256))
    except ValueError:
        return 1


def _escape_markdown_cell(value: str) -> str:
    return html.escape(value, quote=False).replace("\\", "\\\\").replace("|", "\\|").replace("\n", "<br>")


def _html_table_to_markdown(table_html: str, *, minimum_matches: int = 2, merge_strategy: str = "fill") -> str:
    parser = _TableHTMLParser()
    parser.feed(table_html)
    cells: dict[tuple[int, int], str] = {}
    occupied: set[tuple[int, int]] = set()
    merges: list[TableMergeRegion] = []
    height, width = len(parser.rows), 0
    for row, values in enumerate(parser.rows):
        col = 0
        for value, rowspan, colspan in values:
            while (row, col) in occupied:
                col += 1
            covered = {(r, c) for r in range(row, row + rowspan) for c in range(col, col + colspan)}
            if covered & occupied:
                return ""
            occupied.update(covered)
            cells[row, col] = value
            if rowspan > 1 or colspan > 1:
                merges.append(TableMergeRegion(row, col, row + rowspan - 1, col + colspan - 1))
            height, width = max(height, row + rowspan), max(width, col + colspan)
            col += colspan
    if height < 2 or width < 2 or height * width > 65536 or sum(bool(v) for v in cells.values()) < minimum_matches:
        return ""
    grid = render_table_semantic_grid(
        build_table_semantic_grid(row_count=height, col_count=width, cell_text_by_position=cells, merge_regions=merges),
        strategy=merge_strategy,
    )
    lines = ["| " + " | ".join(_escape_markdown_cell(cell) for cell in row) + " |" for row in grid]
    lines.insert(1, "| " + " | ".join("---" for _ in range(width)) + " |")
    return "\n".join(lines)


@lru_cache(maxsize=32)
def _model_digest(path: str, size: int, modified: int, changed: int) -> str:
    del size, modified, changed
    with Path(path).open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def _resolve_model(filename: str, digest: str, override: str, explicit: str | Path | None = None) -> Path | None:
    selected = explicit or os.environ.get(override)
    if selected:
        candidates = [Path(selected)]
    else:
        roots = list(Path(__file__).resolve().parents)
        resource = os.environ.get("DOCWEN_RESOURCE_ROOT")
        if resource:
            roots.insert(0, Path(resource))
        frozen = getattr(sys, "_MEIPASS", None)
        if frozen:
            roots.insert(0, Path(frozen))
        roots.append(Path(sys.executable).resolve().parent)
        candidates = [root / "models" / "rapidtable" / filename for root in roots]
    for candidate in candidates:
        try:
            stat = candidate.stat()
            if (
                candidate.is_file()
                and _model_digest(str(candidate.resolve()), stat.st_size, stat.st_mtime_ns, stat.st_ctime_ns) == digest
            ):
                return candidate
        except OSError:
            continue
    return None


def resolve_rapidtable_model(model_path: str | Path | None = None) -> Path | None:
    return _resolve_model(RAPIDTABLE_MODEL_FILENAME, RAPIDTABLE_MODEL_SHA256, "DOCWEN_RAPIDTABLE_MODEL", model_path)


def resolve_table_layout_model(model_path: str | Path | None = None) -> Path | None:
    return _resolve_model(
        TABLE_LAYOUT_MODEL_FILENAME, TABLE_LAYOUT_MODEL_SHA256, "DOCWEN_TABLE_LAYOUT_MODEL", model_path
    )


def table_recognition_available(model_path: str | Path | None = None) -> bool:
    return resolve_rapidtable_model(model_path) is not None and resolve_table_layout_model() is not None


def _request_table_options(context: object) -> tuple[bool, str, Callable[[], None] | None]:
    request = getattr(context, "request", None)
    options = getattr(request, "options", {})
    options = options if isinstance(options, Mapping) else {}
    snapshot = getattr(request, "config_snapshot", {})
    ocr = snapshot.get("ocr", {}) if isinstance(snapshot, Mapping) else {}
    default = ocr.get("recognize_tables", True) if isinstance(ocr, Mapping) else True
    enabled = options.get("recognize_tables", default)
    if type(enabled) is not bool:
        raise ValueError("recognize_tables must be a boolean")
    from docwen_core.export_semantics import MarkdownExportSemantics, normalize_table_merge_export_strategy

    semantics = getattr(context, "markdown_export_semantics", None)
    if not isinstance(semantics, MarkdownExportSemantics):
        semantics = MarkdownExportSemantics.from_config_snapshot(snapshot if isinstance(snapshot, Mapping) else {})
    strategy = normalize_table_merge_export_strategy(
        options.get("table_merge_strategy"), default_strategy=semantics.table_merge_export_strategy
    )
    check = getattr(getattr(context, "cancellation", None), "check", None)
    return enabled, strategy, cast("Callable[[], None]", check) if callable(check) else None


def enrich_ocr_table_structure(
    image_path: str | Path,
    outcome: OcrOutcome,
    *,
    enabled: bool = True,
    context: object | None = None,
    merge_strategy: str = "fill",
    check_cancelled: Callable[[], None] | None = None,
) -> tuple[OcrOutcome, TableRecognitionOutcome]:
    table = recognize_table_markdown(
        image_path,
        outcome,
        enabled=enabled,
        context=context,
        merge_strategy=merge_strategy,
        check_cancelled=check_cancelled,
    )
    return (replace(outcome, structured_markdown=table.markdown) if table.recognized_table else outcome), table


def recognize_table_markdown(
    image_path: str | Path,
    ocr_outcome: OcrOutcome,
    *,
    model_path: str | Path | None = None,
    layout_model_path: str | Path | None = None,
    enabled: bool = True,
    context: object | None = None,
    merge_strategy: str = "fill",
    check_cancelled: Callable[[], None] | None = None,
) -> TableRecognitionOutcome:
    if context is not None:
        enabled, merge_strategy, check_cancelled = _request_table_options(context)
    if not enabled:
        return TableRecognitionOutcome(TableRecognitionStatus.DISABLED)
    if ocr_outcome.status is not OcrStatus.SUCCESS:
        return TableRecognitionOutcome(TableRecognitionStatus.NOT_TABLE)
    regions = ocr_outcome.regions
    if len(regions) < 2 or any(len(r.points) < 4 for r in regions):
        return TableRecognitionOutcome(TableRecognitionStatus.NOT_TABLE)
    model, layout = resolve_rapidtable_model(model_path), resolve_table_layout_model(layout_model_path)
    if model is None or layout is None:
        return TableRecognitionOutcome(
            TableRecognitionStatus.UNAVAILABLE, message="Table recognition model is missing or corrupt"
        )
    check = check_cancelled or (lambda: None)
    try:
        import numpy as np

        from docwen_core.text._slanet_table import _load_image, infer_table_structure
        from docwen_core.text._table_layout import (
            add_borderless_candidates,
            detect_table_regions,
            refine_ruled_table_bounds,
        )

        check()
        image = _load_image(Path(image_path))
        rectangles = refine_ruled_table_bounds(image, detect_table_regions(image, layout))
        check()
        boxes = np.asarray([r.points for r in regions], dtype=np.float32)
        rectangles = add_borderless_candidates(rectangles, boxes, image.shape[:2])
        centers = (boxes.min(axis=1) + boxes.max(axis=1)) / 2
        consumed: set[int] = set()
        blocks: list[tuple[float, float, str]] = []
        failures = 0
        for left, top, right, bottom in sorted(rectangles, key=lambda rect: (rect[1], rect[0])):
            check()
            indexes = [
                i for i, (x, y) in enumerate(centers) if left <= x < right and top <= y < bottom and i not in consumed
            ]
            if len(indexes) < 2 or right <= left or bottom <= top:
                continue
            try:
                raw, matched = infer_table_structure(
                    image[top:bottom, left:right],
                    model,
                    boxes=boxes[indexes] - [left, top],
                    texts=tuple(regions[i].text for i in indexes),
                    scores=tuple(regions[i].confidence for i in indexes),
                )
                check()
                markdown = _html_table_to_markdown(raw, merge_strategy=merge_strategy)
                if not markdown or not matched or any(i < 0 or i >= len(indexes) for i in matched):
                    failures += 1
                    continue
                consumed.update(indexes[i] for i in matched)
                blocks.append((float(min(indexes[i] for i in matched)), 0.0, markdown))
            except CancellationRequested:
                raise
            except Exception:
                failures += 1
        count = len(blocks)
        if not count:
            return TableRecognitionOutcome(
                TableRecognitionStatus.FAILED if failures else TableRecognitionStatus.NOT_TABLE,
                message="Table structure was not usable; plain OCR text retained" if failures else "",
            )
        blocks.extend(
            (float(i), 0.0, html.escape(r.text, quote=False))
            for i, r in enumerate(regions)
            if i not in consumed and r.text
        )
        check()
        return TableRecognitionOutcome(
            TableRecognitionStatus.SUCCESS,
            markdown="\n\n".join(value for _, _, value in sorted(blocks, key=lambda item: item[:2])),
            message="Some table regions fell back to plain OCR" if failures else "",
            table_count=count,
        )
    except CancellationRequested:
        raise
    except Exception as exc:
        return TableRecognitionOutcome(TableRecognitionStatus.FAILED, message=str(exc))
