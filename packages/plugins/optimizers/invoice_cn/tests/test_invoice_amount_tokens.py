"""Exact signed decimal expectations supplement native-file regressions."""

from __future__ import annotations

import pytest

from docwen_plugin_optimizer_invoice_cn.invoice_cn.ocr_normalize import _normalize_ocr_amount

pytestmark = pytest.mark.unit


@pytest.mark.parametrize(
    "source,expected",
    [
        ("￥-100.00", "-100.00"),
        ("-￥100.00", "-100.00"),
        ("−13.0", "-13.00"),
        ("－113", "-113.00"),
        ("+100", "100.00"),
        ("0", "0.00"),
        ("-0.00", "-0.00"),
        ("9,007,199,254,740,993.123456789", "9007199254740993.123456789"),
        ("-1OO.OO", "-100.00"),
        ("1.2.3", ""),
        ("1-00", ""),
        ("1¥2", ""),
        ("NaN", ""),
        ("Infinity", ""),
        ("1,00", ""),
    ],
)
def test_normalized_amount_keeps_exact_sign_precision_and_refuses_invalid_syntax(source: str, expected: str) -> None:
    assert _normalize_ocr_amount(source) == expected
