"""Regression tests for positional PDF invoice metadata extraction."""

from __future__ import annotations

import pytest

pytestmark = pytest.mark.unit

from docwen_plugin_optimizer_invoice_cn.invoice_cn.metadata import (
    parse_invoice_metadata_from_pdf_spans,
)


def _span(x0: float, y0: float, x1: float, text: str) -> tuple[float, float, float, float, str]:
    return (x0, y0, x1, y0 + 8.0, text)


def test_currency_prefixed_grand_total_is_selected_from_total_line() -> None:
    spans = [
        _span(55.0, 275.0, 130.0, "价税合计（小写）"),
        _span(420.0, 256.0, 455.0, "¥100.00"),
        _span(570.0, 256.0, 600.0, "¥3.00"),
        _span(455.0, 275.5, 500.0, "¥103.00"),
    ]

    metadata = parse_invoice_metadata_from_pdf_spans(spans)

    assert metadata["价税合计"] == "103.00"


def test_split_tax_id_labels_infer_distinct_buyer_and_seller_columns() -> None:
    buyer_tax_id = "91330100MA1234567A"
    seller_tax_id = "91330100MA7654321B"
    spans = [
        _span(40.0, 120.0, 115.0, "统一社会信用代码"),
        _span(120.0, 120.0, 175.0, "/纳税人识别号"),
        _span(180.0, 120.0, 290.0, buyer_tax_id),
        _span(325.0, 120.0, 400.0, "统一社会信用代码"),
        _span(405.0, 120.0, 460.0, "/纳税人识别号"),
        _span(465.0, 120.0, 575.0, seller_tax_id),
    ]

    metadata = parse_invoice_metadata_from_pdf_spans(spans)

    assert metadata["购买方纳税人识别号"] == buyer_tax_id
    assert metadata["销售方纳税人识别号"] == seller_tax_id
