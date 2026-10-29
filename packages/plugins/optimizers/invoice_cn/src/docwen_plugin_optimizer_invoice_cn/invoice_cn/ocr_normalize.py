"""Text normalisation helpers for OCR-like corrections in invoice text.

These handle common OCR/copy-paste artifacts in Chinese invoice text:
digits replaced by look-alike letters (O→0, l→1, etc.),
Chinese/English separator characters, and whitespace compaction.
"""

from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation


def _compact_text(text: str) -> str:
    """Remove all whitespace and zero-width characters from text."""
    s = re.sub(r"\s+", "", text or "")
    return s.replace("﻿", "").replace("​", "").replace("‌", "").replace("‍", "")


def _regex_first(text: str, patterns: list[str]) -> str:
    """Return the first capture group match from a list of regex patterns."""
    for p in patterns:
        m = re.search(p, text, flags=re.MULTILINE)
        if m:
            return (m.group(1) or "").strip()
    return ""


def _normalize_ocr_digits(value: str) -> str:
    """Normalise OCR-misread digits: O→0, l→1, Z→2, S→5, B→8, etc."""
    s = (value or "").strip()
    if not s:
        return ""
    table = str.maketrans({"O": "0", "o": "0", "I": "1", "l": "1", "Z": "2", "z": "2", "S": "5", "s": "5", "B": "8"})
    s = s.translate(table)
    s = re.sub(r"[^0-9]", "", s)
    return s


def _normalize_ocr_amount(value: str) -> str:
    """Return a signed decimal string without binary rounding or guessing.

    Known OCR digit/punctuation substitutions are allowed. A sign belongs to
    the amount, and unsupported leftover syntax cannot be silently deleted.
    Monetary fields have at least two decimal places; additional authored
    precision is retained rather than rounded.
    """
    s = (value or "").strip()
    if not s:
        return ""
    s = s.translate(
        str.maketrans(
            {
                "O": "0",
                "o": "0",
                "I": "1",
                "l": "1",
                "Z": "2",
                "z": "2",
                "S": "5",
                "s": "5",
                "B": "8",
                "，": ",",
                "．": ".",
                "。": ".",
                "−": "-",
                "－": "-",
                "＋": "+",
            }
        )
    )
    s = re.sub(r"\s+", "", s)
    if s.startswith(("¥", "￥")):
        s = s[1:]
    elif s[:1] in {"+", "-"} and s[1:2] in {"¥", "￥"}:
        s = s[:1] + s[2:]
    if re.fullmatch(r"[+-]?(?:[0-9]+|[0-9]{1,3}(?:,[0-9]{3})+)(?:\.[0-9]*)?", s) is None:
        return ""
    try:
        amount = Decimal(s.replace(",", ""))
    except InvalidOperation:
        return ""
    if not amount.is_finite():
        return ""
    result = format(amount, "f")
    if "." not in result:
        return result + ".00"
    fraction = result.partition(".")[2]
    return result + "0" * max(0, 2 - len(fraction))


def _normalize_ocr_tax_id(value: str) -> str:
    """Normalise OCR-misread tax ID numbers (letters and digits only)."""
    s = (value or "").strip()
    if not s:
        return ""
    table = str.maketrans({"O": "0", "o": "0", "I": "1", "l": "1", "Z": "2", "z": "2", "S": "5", "s": "5"})
    s = s.translate(table).upper()
    s = re.sub(r"[^0-9A-Z]", "", s)
    return s


def _normalize_ocr_date(value: str) -> str:
    """Normalise OCR-misread dates to standard 'YYYY年MM月DD日' format."""
    s = (value or "").strip()
    if not s:
        return ""
    table = str.maketrans({"O": "0", "o": "0", "I": "1", "l": "1", "Z": "2", "z": "2", "S": "5", "s": "5", "B": "8"})
    s_norm = s.translate(table)
    m = re.search(r"(20[0-9]{2})[年./·•-]([0-9]{1,2})[月./·•-]([0-9]{1,2})", s_norm)
    if not m:
        return s_norm.strip()
    y = m.group(1)
    mm = int(m.group(2))
    dd = int(m.group(3))
    return f"{y}年{mm:02d}月{dd:02d}日"
