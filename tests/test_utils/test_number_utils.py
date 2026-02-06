"""
number_utils 单元测试

覆盖所有数字转换函数的正常值、边界值和超范围情况。
"""

from __future__ import annotations

import pytest

from docwen.utils.number_utils import (
    number_to_chinese,
    number_to_chinese_upper,
    number_to_circled,
    number_to_arabic_full,
    number_to_letter_upper,
    number_to_letter_lower,
    number_to_roman_upper,
    number_to_roman_lower,
)


pytestmark = pytest.mark.unit


# ============================================================
# number_to_chinese
# ============================================================

class TestNumberToChinese:
    """小写中文数字转换"""

    @pytest.mark.parametrize(
        ("num", "expected"),
        [
            (1, "一"),
            (5, "五"),
            (9, "九"),
            (10, "十"),
            (11, "十一"),
            (19, "十九"),
            (20, "二十"),
            (21, "二十一"),
            (55, "五十五"),
            (99, "九十九"),
        ],
    )
    def test_normal_range(self, num: int, expected: str) -> None:
        assert number_to_chinese(num) == expected

    @pytest.mark.parametrize("num", [0, -1, -100])
    def test_non_positive_returns_empty(self, num: int) -> None:
        assert number_to_chinese(num) == ""

    def test_over_99_returns_arabic(self) -> None:
        assert number_to_chinese(100) == "100"
        assert number_to_chinese(999) == "999"


# ============================================================
# number_to_chinese_upper
# ============================================================

class TestNumberToChineseUpper:
    """大写中文数字（财务用）转换"""

    @pytest.mark.parametrize(
        ("num", "expected"),
        [
            (1, "壹"),
            (5, "伍"),
            (10, "拾"),
            (11, "拾壹"),
            (20, "贰拾"),
            (23, "贰拾叁"),
            (99, "玖拾玖"),
        ],
    )
    def test_normal_range(self, num: int, expected: str) -> None:
        assert number_to_chinese_upper(num) == expected

    def test_non_positive_returns_empty(self) -> None:
        assert number_to_chinese_upper(0) == ""
        assert number_to_chinese_upper(-5) == ""

    def test_over_99_returns_arabic(self) -> None:
        assert number_to_chinese_upper(100) == "100"


# ============================================================
# number_to_circled
# ============================================================

class TestNumberToCircled:
    """带圈数字转换"""

    @pytest.mark.parametrize(
        ("num", "expected"),
        [
            (1, "①"),
            (10, "⑩"),
            (20, "⑳"),
            (50, "㊿"),
        ],
    )
    def test_normal_range(self, num: int, expected: str) -> None:
        assert number_to_circled(num) == expected

    @pytest.mark.parametrize("num", [0, -1, 51, 100])
    def test_out_of_range_returns_parenthesized(self, num: int) -> None:
        assert number_to_circled(num) == f"({num})"


# ============================================================
# number_to_arabic_full
# ============================================================

class TestNumberToArabicFull:
    """全角阿拉伯数字转换"""

    @pytest.mark.parametrize(
        ("num", "expected"),
        [
            (0, "０"),
            (1, "１"),
            (10, "１０"),
            (99, "９９"),
            (123, "１２３"),
        ],
    )
    def test_conversions(self, num: int, expected: str) -> None:
        assert number_to_arabic_full(num) == expected


# ============================================================
# number_to_letter_upper / lower
# ============================================================

class TestNumberToLetter:
    """拉丁字母转换"""

    @pytest.mark.parametrize(
        ("num", "expected"),
        [
            (1, "A"),
            (26, "Z"),
            (27, "AA"),
            (28, "AB"),
            (52, "AZ"),
            (53, "BA"),
        ],
    )
    def test_upper(self, num: int, expected: str) -> None:
        assert number_to_letter_upper(num) == expected

    @pytest.mark.parametrize(
        ("num", "expected"),
        [
            (1, "a"),
            (26, "z"),
            (27, "aa"),
        ],
    )
    def test_lower(self, num: int, expected: str) -> None:
        assert number_to_letter_lower(num) == expected

    @pytest.mark.parametrize("num", [0, -1])
    def test_non_positive_returns_empty(self, num: int) -> None:
        assert number_to_letter_upper(num) == ""
        assert number_to_letter_lower(num) == ""


# ============================================================
# number_to_roman_upper / lower
# ============================================================

class TestNumberToRoman:
    """罗马数字转换"""

    @pytest.mark.parametrize(
        ("num", "expected"),
        [
            (1, "I"),
            (4, "IV"),
            (9, "IX"),
            (14, "XIV"),
            (40, "XL"),
            (90, "XC"),
            (400, "CD"),
            (900, "CM"),
            (1994, "MCMXCIV"),
            (3999, "MMMCMXCIX"),
        ],
    )
    def test_upper(self, num: int, expected: str) -> None:
        assert number_to_roman_upper(num) == expected

    @pytest.mark.parametrize(
        ("num", "expected"),
        [
            (1, "i"),
            (4, "iv"),
            (1994, "mcmxciv"),
        ],
    )
    def test_lower(self, num: int, expected: str) -> None:
        assert number_to_roman_lower(num) == expected

    @pytest.mark.parametrize("num", [0, -1, 4000, 5000])
    def test_out_of_range_returns_arabic(self, num: int) -> None:
        assert number_to_roman_upper(num) == str(num)
        assert number_to_roman_lower(num) == str(num).lower()
