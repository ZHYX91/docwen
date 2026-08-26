"""Number-to-text conversion utilities — Chinese, Latin, Roman, circled, full-width.

Provides a single source of truth for all numbering styles used across
plugins (document, layout, image, spreadsheet).  Every function is pure,
has no side effects, and carries zero dependencies beyond the standard
library.

Audit mapping
  F-I2b-011  ``number_to_chinese_upper`` — Chinese uppercase (financial) numerals
  F-I2b-014  ``number_to_letter_upper``    — Latin uppercase A, B, …, Z, AA, AB…
  (F-I2b-010 through F-I2b-019 are covered by the full function set below)
"""

from __future__ import annotations

# ──────────────────────────────────────────────────────────────────────────────
# Chinese numerals — lowercase (一二三…) and uppercase 壹贰叁 (financial)
# ──────────────────────────────────────────────────────────────────────────────

_LOWER_DIGITS = ["", "一", "二", "三", "四", "五", "六", "七", "八", "九"]
_UPPER_DIGITS = ["", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"]


def _chinese_2digit(n: int, digits: list[str], ten_label: str) -> str:
    """Render 1 <= n < 100 as Chinese numeral using *digits* table."""
    if n <= 10:
        return digits[n] if n <= 9 else ten_label
    if n < 20:
        return ten_label + digits[n - 10]
    tens = digits[n // 10] + ten_label
    ones = n % 10
    return tens if ones == 0 else tens + digits[ones]


def number_to_chinese(num: int) -> str:
    """Convert an integer (1–99) to lowercase Chinese numeral.

    >>> number_to_chinese(1)
    '一'
    >>> number_to_chinese(12)
    '十二'
    >>> number_to_chinese(100)
    '100'
    """
    if num <= 0:
        return ""
    if num >= 100:
        return str(num)
    return _chinese_2digit(num, _LOWER_DIGITS, "十")


def number_to_chinese_upper(num: int) -> str:
    """Convert an integer (1–99) to uppercase (financial) Chinese numeral.

    >>> number_to_chinese_upper(1)
    '壹'
    >>> number_to_chinese_upper(23)
    '贰拾叁'
    >>> number_to_chinese_upper(100)
    '100'
    """
    if num <= 0:
        return ""
    if num >= 100:
        return str(num)
    return _chinese_2digit(num, _UPPER_DIGITS, "拾")


# ──────────────────────────────────────────────────────────────────────────────
# Circled numbers  ① … ⑳  ㉑ … ㊿
# ──────────────────────────────────────────────────────────────────────────────

_CIRCLED: list[str] = [
    "①",
    "②",
    "③",
    "④",
    "⑤",
    "⑥",
    "⑦",
    "⑧",
    "⑨",
    "⑩",
    "⑪",
    "⑫",
    "⑬",
    "⑭",
    "⑮",
    "⑯",
    "⑰",
    "⑱",
    "⑲",
    "⑳",
    "㉑",
    "㉒",
    "㉓",
    "㉔",
    "㉕",
    "㉖",
    "㉗",
    "㉘",
    "㉙",
    "㉚",
    "㉛",
    "㉜",
    "㉝",
    "㉞",
    "㉟",
    "㊱",
    "㊲",
    "㊳",
    "㊴",
    "㊵",
    "㊶",
    "㊷",
    "㊸",
    "㊹",
    "㊺",
    "㊻",
    "㊼",
    "㊽",
    "㊾",
    "㊿",
]


def number_to_circled(num: int) -> str:
    """Convert an integer (1–50) to a circled Unicode character.

    >>> number_to_circled(1)
    '①'
    >>> number_to_circled(50)
    '㊿'
    >>> number_to_circled(51)
    '(51)'
    """
    if 1 <= num <= 50:
        return _CIRCLED[num - 1]
    return f"({num})"


# ──────────────────────────────────────────────────────────────────────────────
# Full-width Arabic digits  ０１２… → １２３
# ──────────────────────────────────────────────────────────────────────────────

_FULLWIDTH_MAP = {str(i): c for i, c in enumerate("０１２３４５６７８９")}


def number_to_arabic_full(num: int) -> str:
    """Convert an integer to full-width Arabic numeral string.

    >>> number_to_arabic_full(0)
    '０'
    >>> number_to_arabic_full(99)
    '９９'
    """
    return "".join(_FULLWIDTH_MAP.get(c, c) for c in str(num))


# ──────────────────────────────────────────────────────────────────────────────
# Latin letter sequences  A, B, …, Z, AA, AB, …
# ──────────────────────────────────────────────────────────────────────────────


def _to_letters(num: int, base: int) -> str:
    """Bijective base-26 letter sequence (1 → A, 26 → Z, 27 → AA)."""
    result: list[str] = []
    while num > 0:
        num -= 1
        result.append(chr(base + (num % 26)))
        num //= 26
    return "".join(reversed(result))


def number_to_letter_upper(num: int) -> str:
    """Convert an integer to uppercase Latin letters (bijective base-26).

    >>> number_to_letter_upper(1)
    'A'
    >>> number_to_letter_upper(26)
    'Z'
    >>> number_to_letter_upper(27)
    'AA'
    """
    if num <= 0:
        return ""
    return _to_letters(num, 65)  # ord('A')


def number_to_letter_lower(num: int) -> str:
    """Convert an integer to lowercase Latin letters (bijective base-26).

    >>> number_to_letter_lower(1)
    'a'
    >>> number_to_letter_lower(27)
    'aa'
    """
    if num <= 0:
        return ""
    return _to_letters(num, 97)  # ord('a')


# ──────────────────────────────────────────────────────────────────────────────
# Roman numerals  I, II, …, IV, …, IX, …, MMMCMXCIX (3999)
# ──────────────────────────────────────────────────────────────────────────────

_ROMAN_VALUES = [
    (1000, "M"),
    (900, "CM"),
    (500, "D"),
    (400, "CD"),
    (100, "C"),
    (90, "XC"),
    (50, "L"),
    (40, "XL"),
    (10, "X"),
    (9, "IX"),
    (5, "V"),
    (4, "IV"),
    (1, "I"),
]


def number_to_roman_upper(num: int) -> str:
    """Convert an integer (1–3999) to uppercase Roman numeral.

    >>> number_to_roman_upper(1)
    'I'
    >>> number_to_roman_upper(1994)
    'MCMXCIV'
    >>> number_to_roman_upper(4000)
    '4000'
    """
    if num <= 0 or num >= 4000:
        return str(num)
    parts: list[str] = []
    remainder = num
    for value, symbol in _ROMAN_VALUES:
        count, remainder = divmod(remainder, value)
        if count:
            parts.append(symbol * count)
        if remainder == 0:
            break
    return "".join(parts)


def number_to_roman_lower(num: int) -> str:
    """Convert an integer (1–3999) to lowercase Roman numeral.

    >>> number_to_roman_lower(1)
    'i'
    >>> number_to_roman_lower(4)
    'iv'
    >>> number_to_roman_lower(4000)
    '4000'
    """
    if num <= 0 or num >= 4000:
        return str(num).lower()
    parts: list[str] = []
    remainder = num
    for value, symbol in _ROMAN_VALUES:
        count, remainder = divmod(remainder, value)
        if count:
            parts.append(symbol.lower() * count)
        if remainder == 0:
            break
    return "".join(parts)
