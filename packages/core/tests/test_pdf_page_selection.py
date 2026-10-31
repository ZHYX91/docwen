"""Independent bounded grammar and typed-page expectations."""

from __future__ import annotations

from typing import Any

import pytest

import docwen_core.page_selection as selection_module
from docwen_core.page_selection import (
    MAX_PDF_SPLIT_PAGES,
    PageSelectionError,
    parse_pdf_split_input,
    validate_pdf_split_selection,
)

pytestmark = pytest.mark.unit


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("3,1,3", (1, 3)),
        ("5-3,1", (1, 3, 4, 5)),
        ("1，3；5、7;9", (1, 3, 5, 7, 9)),
        ("5 至 3，1~2,2－4", (1, 2, 3, 4, 5)),
        (" 1 - 3 , 7 ", (1, 2, 3, 7)),
        ("00001,3", (1, 3)),
    ],
)
def test_literal_expected_selection(text: str, expected: tuple[int, ...]) -> None:
    selected = parse_pdf_split_input(text, total_pages=12)
    assert selected.mode == "custom"
    assert selected.pages == expected


@pytest.mark.parametrize(
    ("text", "total", "code"),
    [
        ("1,3x", 4, "invalid_token"),
        ("1,,3", 4, "invalid_token"),
        ("1,", 4, "invalid_token"),
        ("1-3-4", 4, "invalid_token"),
        ("1 3", 4, "invalid_token"),
        ("-1,1", 4, "invalid_token"),
        ("0,1", 4, "out_of_range"),
        ("1,999", 4, "out_of_range"),
        ("1-100000000000000", 4, "out_of_range"),
        ("1-100000,3x", 0, "invalid_token"),
        ("1-4", 4, "all_pages"),
        ("4-1,3,2", 4, "all_pages"),
        ("1," * 1024 + "1", 4, "token_budget"),
        ("1" * 4097, 4, "text_budget"),
    ],
)
def test_every_failure_is_proven_before_any_range_is_materialized(
    monkeypatch: pytest.MonkeyPatch, text: str, total: int, code: str
) -> None:
    def forbidden_expansion(*_args: object) -> Any:
        raise AssertionError("an invalid complete selection reached range expansion")

    # This guards the allocation boundary, not a mock call-count/content oracle.
    # Separate file tests verify real CLI/Runtime publication and exact PDF pages.
    monkeypatch.setattr(selection_module, "range", forbidden_expansion, raising=False)
    with pytest.raises(PageSelectionError) as caught:
        parse_pdf_split_input(text, total_pages=total)
    assert caught.value.code == code


@pytest.mark.parametrize("total", [True, False, 4.0, float("nan"), float("inf"), "4", -1, MAX_PDF_SPLIT_PAGES + 1])
def test_page_count_is_an_exact_bounded_integer(total: Any) -> None:
    with pytest.raises(PageSelectionError):
        parse_pdf_split_input("1", total_pages=total)
    with pytest.raises(PageSelectionError):
        validate_pdf_split_selection("custom", [1], total_pages=total)


@pytest.mark.parametrize("values", [[1, True], [1, 3.0], [1, float("nan")], [1, "3"], [1, 999], None, "1,3"])
def test_runtime_does_not_filter_or_coerce_typed_values(values: object) -> None:
    with pytest.raises(PageSelectionError):
        validate_pdf_split_selection("custom", values, total_pages=4)


def test_overlap_union_budget_and_runtime_recheck_have_independent_expectations() -> None:
    selected = parse_pdf_split_input("1-100000,1-100000")
    assert len(selected.pages) == MAX_PDF_SPLIT_PAGES
    assert selected.pages[0] == 1 and selected.pages[-1] == 100000
    with pytest.raises(PageSelectionError) as caught:
        validate_pdf_split_selection("custom", list(selected.pages), total_pages=4)
    assert caught.value.code == "out_of_range"
    assert validate_pdf_split_selection("custom", [3, 1, 3], total_pages=4).pages == (1, 3)


@pytest.mark.parametrize(("text", "mode"), [("*", "every_page"), ("#", "odd_even")])
def test_wildcards_have_modes_without_custom_full_selection(text: str, mode: str) -> None:
    selected = parse_pdf_split_input(text, total_pages=4)
    assert (selected.mode, selected.pages) == (mode, ())
    with pytest.raises(PageSelectionError) as caught:
        parse_pdf_split_input(text, total_pages=1)
    assert caught.value.code == "single_page"
