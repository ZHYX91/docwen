"""utils 单元测试。"""

from __future__ import annotations

import pytest

from docwen.utils.validation_utils import (
    contains_chinese,
    is_chinese,
    is_value_empty,
    validate_date_format,
    validate_ocr_requires_images,
)


pytestmark = pytest.mark.unit


@pytest.mark.unit
@pytest.mark.parametrize(
    ("value", "expected"),
    [
        (None, True),
        ("", True),
        ("  ", True),
        ([], True),
        (["", None], True),
        (["", "x"], False),
        ({}, True),
        ({"a": ""}, True),
        ({"a": "x"}, False),
        ("内容", False),
        (0, False),
    ],
)
def test_is_value_empty(value, expected: bool) -> None:
    assert is_value_empty(value) is expected


@pytest.mark.unit
def test_is_chinese_and_contains_chinese() -> None:
    assert is_chinese("文") is True
    assert is_chinese("A") is False
    assert contains_chinese("ABC文DEF") is True
    assert contains_chinese("ABC") is False


@pytest.mark.unit
@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("2023年12月31日", True),
        ("2023-12-31", True),
        ("2023/12/31", True),
        ("2023.12.31", True),
        ("31/12/2023", False),
        ("2023-1-1", False),
    ],
)
def test_validate_date_format(raw: str, expected: bool) -> None:
    assert validate_date_format(raw) is expected


@pytest.mark.unit
@pytest.mark.parametrize(
    ("extract_images", "extract_ocr", "ok"),
    [
        (True, True, True),
        (True, False, True),
        (False, False, True),
        (False, True, True),
    ],
)
def test_validate_ocr_requires_images(extract_images: bool, extract_ocr: bool, ok: bool) -> None:
    actual_ok, reason = validate_ocr_requires_images(extract_images, extract_ocr)
    assert actual_ok is ok
    if ok:
        assert reason == ""
    else:
        assert reason
