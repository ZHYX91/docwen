"""Synthetic saved text PDFs exercise complete signed amount ownership.

The expected amounts and eight-column rows are authored facts, independent of
the compact extractor and the normalizer. These are not OCR or original-report
fixtures.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from tests.integration.test_invoice_native_field_identity import _page_items, _pdf, _run

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]


@pytest.mark.parametrize(
    "token,amount,tax,total,quantity,unit_price",
    [
        ("￥-1,000.00", "-900.00", "-100.00", "-1000.00", "-1", "900.00"),
        ("-￥1,000.00", "-900.00", "-100.00", "-1000.00", "-1", "900.00"),
        ("-￥113.00", "-100.00", "-13.00", "-113.00", "-1", "100.00"),
        ("￥-113.00", "-100.00", "-13.00", "-113.00", "-1", "100.00"),
        ("￥1,000.00", "900.00", "100.00", "1000.00", "1", "900.00"),
        ("+￥1,000.00", "900.00", "100.00", "1000.00", "1", "900.00"),
        ("￥－1，000．00", "-900.00", "-100.00", "-1000.00", "-1", "900.00"),
        ("－￥1，000。00", "-900.00", "-100.00", "-1000.00", "-1", "900.00"),
        ("￥0.00", "0.00", "0.00", "0.00", "0", "100.00"),
        ("￥1.", "1.00", "0.00", "1.00", "1", "1.00"),
    ],
)
@pytest.mark.parametrize("following_label", [False, True], ids=["end", "next-field"])
def test_complete_amount_token_survives_pdf_runtime_publication(
    tmp_path: Path,
    token: str,
    amount: str,
    tax: str,
    total: str,
    quantity: str,
    unit_price: str,
    following_label: bool,
) -> None:
    items, row = _page_items(amount=amount, tax=tax, total=total, quantity=quantity, unit_price=unit_price)
    items[-1] = (560.0, 400.0, token)
    if following_label:
        items.append((40.0, 430.0, "开票人：张三"))
    source = tmp_path / "complete-amount.pdf"
    _pdf(source, [items])
    metadata, rows, result = _run(source, tmp_path)
    assert {key: metadata[key] for key in ("金额", "税额", "价税合计")} == {
        "金额": amount,
        "税额": tax,
        "价税合计": total,
    }
    assert all(isinstance(metadata[key], str) for key in ("金额", "税额", "价税合计"))
    assert rows == [row]
    assert not any(d.code == "INVOICE-METADATA-CONFLICT" for d in result.diagnostics)


@pytest.mark.parametrize(
    "token",
    [
        "￥-1,00.00",
        "￥-12,34,567.00",
        "￥-1234,567.00",
        "￥-1,000.00junk",
        "￥-1,000.00尾随数据",
        "￥-113.00%",
        "￥-113.00/7",
        "￥-113.00_7",
        "prefix￥-113.00",
        "￥--113.00",
        "￥-113.00.4",
        "￥-113.00-7",
    ],
)
def test_invalid_amount_is_not_a_valid_numeric_prefix(tmp_path: Path, token: str) -> None:
    items, row = _page_items()
    items[-1] = (560.0, 400.0, token)
    source = tmp_path / "invalid-amount.pdf"
    _pdf(source, [items])
    metadata, rows, _ = _run(source, tmp_path)
    assert metadata["价税合计"] == ""
    assert metadata["金额"] == "-100.00"
    assert metadata["税额"] == "-13.00"
    assert rows == [row]
