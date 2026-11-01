"""Strict, bounded page selection shared by CLI, GUI and PDF execution."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Literal

MAX_PAGE_SELECTION_TEXT = 4096
MAX_PAGE_SELECTION_TOKENS = 1024
MAX_PDF_SPLIT_PAGES = 100_000

_SEPARATORS = str.maketrans({"，": ",", ";": ",", "；": ",", "、": ",", "~": "-", "－": "-", "至": "-"})
_TOKEN = re.compile(r"(\d+)\s*(?:-\s*(\d+))?")


class PageSelectionError(ValueError):
    """A complete selection is invalid; callers must never keep a subset."""

    def __init__(self, code: str, message: str) -> None:
        self.code = code
        super().__init__(message)


@dataclass(frozen=True, slots=True)
class PdfPageSelection:
    mode: Literal["custom", "every_page", "odd_even"]
    pages: tuple[int, ...] = ()


def validate_pdf_page_count(total_pages: object, *, allow_unknown: bool = False) -> int:
    """Validate source metadata without coercing bool, float or non-finite values."""
    if not isinstance(total_pages, int) or isinstance(total_pages, bool):
        raise PageSelectionError("invalid_total", "PDF page count must be an integer.")
    minimum = 0 if allow_unknown else 1
    if total_pages < minimum or total_pages > MAX_PDF_SPLIT_PAGES:
        raise PageSelectionError("page_budget", "PDF page count exceeds the supported split budget.")
    return total_pages


def _check_split_count(total_pages: int) -> None:
    if total_pages == 1:
        raise PageSelectionError("single_page", "PDF has only one page; nothing to split.")


def _check_selection_size(size: int, total_pages: int) -> None:
    if size == 0:
        raise PageSelectionError("empty_selection", "No pages specified for custom split.")
    if size > MAX_PDF_SPLIT_PAGES:
        raise PageSelectionError("page_budget", "Selected pages exceed the supported split budget.")
    if total_pages and size == total_pages:
        raise PageSelectionError("all_pages", "Custom selection covers the entire document.")


def _page_number(value: object, total_pages: int) -> int:
    if not isinstance(value, int) or isinstance(value, bool):
        raise PageSelectionError("invalid_page", "Each page must be an integer.")
    if value < 1 or value > MAX_PDF_SPLIT_PAGES or (total_pages and value > total_pages):
        raise PageSelectionError("out_of_range", "Every selected page must be within the source and split budget.")
    return value


def parse_pdf_split_input(expression: str, *, total_pages: int = 0) -> PdfPageSelection:
    """Validate every token and endpoint before materializing the union of ranges.

    Zero means the GUI/CLI has not read source metadata yet. Execution must
    validate the resulting selection again against the actual PDF page count.
    """
    total = validate_pdf_page_count(total_pages, allow_unknown=True)
    if not isinstance(expression, str) or len(expression) > MAX_PAGE_SELECTION_TEXT:
        raise PageSelectionError("text_budget", "Page expression exceeds the supported input budget.")
    text = expression.strip()
    if text == "*":
        _check_split_count(total)
        return PdfPageSelection("every_page")
    if text == "#":
        _check_split_count(total)
        return PdfPageSelection("odd_even")
    parts = text.translate(_SEPARATORS).split(",")
    if len(parts) > MAX_PAGE_SELECTION_TOKENS:
        raise PageSelectionError("token_budget", "Page expression contains too many tokens.")
    intervals: list[tuple[int, int]] = []
    for part in parts:
        match = _TOKEN.fullmatch(part.strip())
        if match is None:
            raise PageSelectionError("invalid_token", "Every page token must be a positive page or complete range.")
        start = _page_number(int(match[1]), total)
        end = _page_number(int(match[2]), total) if match[2] is not None else start
        intervals.append((min(start, end), max(start, end)))

    # Merging at most MAX_PAGE_SELECTION_TOKENS intervals proves the exact
    # distinct cardinality without allocating any range's individual pages.
    merged: list[tuple[int, int]] = []
    for start, end in sorted(intervals):
        if merged and start <= merged[-1][1] + 1:
            merged[-1] = (merged[-1][0], max(end, merged[-1][1]))
        else:
            merged.append((start, end))
    _check_selection_size(sum(end - start + 1 for start, end in merged), total)
    pages = tuple(page for start, end in merged for page in range(start, end + 1))
    return PdfPageSelection("custom", pages)


def validate_pdf_split_selection(mode: object, pages: object, *, total_pages: int) -> PdfPageSelection:
    """Revalidate typed Runtime options against actual, bounded source metadata."""
    total = validate_pdf_page_count(total_pages)
    _check_split_count(total)
    if mode == "every_page":
        return PdfPageSelection("every_page")
    if mode == "odd_even":
        return PdfPageSelection("odd_even")
    if mode != "custom":
        raise PageSelectionError("invalid_mode", "Unsupported PDF split mode.")
    if not isinstance(pages, (list, tuple)):
        raise PageSelectionError("invalid_pages", "Custom PDF pages must be an array of integers.")
    if len(pages) > MAX_PDF_SPLIT_PAGES:
        raise PageSelectionError("page_budget", "Selected pages exceed the supported split budget.")
    # No filtering/coercion: inspect the complete typed input before deduping.
    checked = [_page_number(value, total) for value in pages]
    selected = tuple(sorted(set(checked)))
    _check_selection_size(len(selected), total)
    return PdfPageSelection("custom", selected)
