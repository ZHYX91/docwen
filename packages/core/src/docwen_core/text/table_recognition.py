"""Offline table-structure recognition over RapidTable.

The runtime never downloads models. Packaged builds place the pinned
SLANet-plus ONNX asset under models/rapidtable/; source/dev installs may
provide the same file through DOCWEN_RAPIDTABLE_MODEL.
"""

from __future__ import annotations

import html
import os
import sys
from dataclasses import dataclass, replace
from enum import StrEnum
from html.parser import HTMLParser
from pathlib import Path

from docwen_core.text.ocr import OcrOutcome, OcrStatus

RAPIDTABLE_MODEL_FILENAME = "slanet-plus.onnx"
RAPIDTABLE_MODEL_SHA256 = "d57a942af6a2f57d6a4a0372573c696a2379bf5857c45e2ac69993f3b334514b"
RAPIDTABLE_MODEL_URL = (
    "https://www.modelscope.cn/models/RapidAI/RapidTable/resolve/v2.0.0/slanet-plus.onnx"
)


class TableRecognitionStatus(StrEnum):
    SUCCESS = "success"
    NOT_TABLE = "not_table"
    UNAVAILABLE = "unavailable"
    FAILED = "failed"


@dataclass(frozen=True, slots=True)
class TableRecognitionOutcome:
    status: TableRecognitionStatus
    markdown: str = ""
    message: str = ""

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
        lowered = tag.lower()
        if lowered == "tr":
            self._row = []
        elif lowered in {"td", "th"} and self._row is not None:
            attr_map = {key.lower(): value for key, value in attrs}
            self._rowspan = _positive_span(attr_map.get("rowspan"))
            self._colspan = _positive_span(attr_map.get("colspan"))
            self._cell_parts = []
        elif lowered == "br" and self._cell_parts is not None:
            self._cell_parts.append("\n")

    def handle_data(self, data: str) -> None:
        if self._cell_parts is not None:
            self._cell_parts.append(data)

    def handle_endtag(self, tag: str) -> None:
        lowered = tag.lower()
        if lowered in {"td", "th"} and self._cell_parts is not None and self._row is not None:
            value = " ".join("".join(self._cell_parts).split())
            self._row.append((html.unescape(value), self._rowspan, self._colspan))
            self._cell_parts = None
            self._rowspan = 1
            self._colspan = 1
        elif lowered == "tr" and self._row is not None:
            if self._row:
                self.rows.append(self._row)
            self._row = None


def _positive_span(value: str | None) -> int:
    try:
        span = int(value or "1")
    except ValueError:
        return 1
    return max(1, min(span, 256))


def _expand_rows(rows: list[list[tuple[str, int, int]]]) -> list[list[str]]:
    grid: list[list[str]] = []
    occupied: dict[tuple[int, int], str] = {}

    for row_index, row in enumerate(rows):
        while len(grid) <= row_index:
            grid.append([])
        column = 0
        for value, rowspan, colspan in row:
            while (row_index, column) in occupied:
                grid[row_index].append(occupied[(row_index, column)])
                column += 1
            for dr in range(rowspan):
                for dc in range(colspan):
                    occupied[(row_index + dr, column + dc)] = value
            for _ in range(colspan):
                grid[row_index].append(value)
            column += colspan

        while (row_index, column) in occupied:
            grid[row_index].append(occupied[(row_index, column)])
            column += 1

    if not grid:
        return []
    width = max(len(row) for row in grid)
    for row_index, row in enumerate(grid):
        for column in range(len(row), width):
            row.append(occupied.get((row_index, column), ""))
    return grid


def _escape_markdown_cell(value: str) -> str:
    return value.replace("\\", "\\\\").replace("|", "\\|").replace("\n", "<br>")


def _html_table_to_markdown(table_html: str, *, minimum_matches: int) -> str:
    parser = _TableHTMLParser()
    parser.feed(table_html)
    grid = _expand_rows(parser.rows)
    if len(grid) < 2 or max((len(row) for row in grid), default=0) < 2:
        return ""

    nonempty = sum(bool(cell.strip()) for row in grid for cell in row)
    if nonempty < max(4, minimum_matches):
        return ""

    width = max(len(row) for row in grid)
    normalized = [row + [""] * (width - len(row)) for row in grid]
    header = normalized[0]
    body = normalized[1:]
    lines = [
        "| " + " | ".join(_escape_markdown_cell(cell) for cell in header) + " |",
        "| " + " | ".join("---" for _ in header) + " |",
    ]
    lines.extend("| " + " | ".join(_escape_markdown_cell(cell) for cell in row) + " |" for row in body)
    return "\n".join(lines)


def _model_candidates(model_path: str | Path | None) -> list[Path]:
    candidates: list[Path] = []
    if model_path is not None:
        candidates.append(Path(model_path))
    env_model = os.environ.get("DOCWEN_RAPIDTABLE_MODEL")
    if env_model:
        candidates.append(Path(env_model))

    roots: list[Path] = []
    env_root = os.environ.get("DOCWEN_RESOURCE_ROOT")
    if env_root:
        roots.append(Path(env_root))
    meipass = getattr(sys, "_MEIPASS", None)
    if isinstance(meipass, str) and meipass:
        roots.extend((Path(meipass), Path(meipass).parent))
    roots.append(Path.cwd())
    roots.extend(Path(__file__).resolve().parents)
    executable = getattr(sys, "executable", "")
    if executable:
        roots.append(Path(executable).resolve().parent)

    for root in roots:
        candidates.extend(
            (
                root / "models" / "rapidtable" / RAPIDTABLE_MODEL_FILENAME,
                root / "rapidtable" / RAPIDTABLE_MODEL_FILENAME,
            )
        )
    return candidates


def resolve_rapidtable_model(model_path: str | Path | None = None) -> Path | None:
    for candidate in _model_candidates(model_path):
        try:
            if candidate.is_file():
                return candidate
        except OSError:
            continue
    return None


def enrich_ocr_table_structure(
    image_path: str | Path,
    outcome: OcrOutcome,
    *,
    model_path: str | Path | None = None,
) -> tuple[OcrOutcome, TableRecognitionOutcome]:
    """Return *outcome* enriched with table Markdown when structure is proven."""

    table_outcome = recognize_table_markdown(image_path, outcome, model_path=model_path)
    if table_outcome.status is TableRecognitionStatus.SUCCESS:
        outcome = replace(outcome, structured_markdown=table_outcome.markdown)
    return outcome, table_outcome


def table_recognition_available(model_path: str | Path | None = None) -> bool:
    return resolve_rapidtable_model(model_path) is not None


def recognize_table_markdown(
    image_path: str | Path,
    ocr_outcome: OcrOutcome,
    *,
    model_path: str | Path | None = None,
) -> TableRecognitionOutcome:
    """Recognize one table-dominant image without performing a second OCR pass."""
    if ocr_outcome.status is not OcrStatus.SUCCESS:
        return TableRecognitionOutcome(TableRecognitionStatus.NOT_TABLE)
    regions = tuple(region for region in ocr_outcome.regions if len(region.points) >= 4)
    if len(regions) < 4:
        return TableRecognitionOutcome(TableRecognitionStatus.NOT_TABLE)

    resolved_model = resolve_rapidtable_model(model_path)
    if resolved_model is None:
        return TableRecognitionOutcome(
            TableRecognitionStatus.UNAVAILABLE,
            message="RapidTable model is unavailable",
        )

    try:
        import numpy as np

        from docwen_core.text._slanet_table import infer_table_html

        boxes = np.asarray([region.points for region in regions], dtype=np.float32)
        texts = tuple(region.text for region in regions)
        scores = tuple(region.confidence for region in regions)
        raw_html = infer_table_html(
            image_path,
            resolved_model,
            boxes=boxes,
            texts=texts,
            scores=scores,
        )
        if not raw_html:
            return TableRecognitionOutcome(TableRecognitionStatus.NOT_TABLE)
        normalized_html = html.unescape(raw_html)
        matched_regions = sum(1 for region in regions if region.text and region.text in normalized_html)
        minimum_matches = min(4, max(2, len(regions) // 3))
        if matched_regions < minimum_matches:
            return TableRecognitionOutcome(TableRecognitionStatus.NOT_TABLE)

        markdown = _html_table_to_markdown(raw_html, minimum_matches=minimum_matches)
        if not markdown:
            return TableRecognitionOutcome(TableRecognitionStatus.NOT_TABLE)
        return TableRecognitionOutcome(TableRecognitionStatus.SUCCESS, markdown=markdown)
    except Exception as exc:
        return TableRecognitionOutcome(TableRecognitionStatus.FAILED, message=str(exc))


__all__ = [
    "RAPIDTABLE_MODEL_FILENAME",
    "RAPIDTABLE_MODEL_SHA256",
    "RAPIDTABLE_MODEL_URL",
    "TableRecognitionOutcome",
    "TableRecognitionStatus",
    "enrich_ocr_table_structure",
    "recognize_table_markdown",
    "resolve_rapidtable_model",
    "table_recognition_available",
]
