"""Regression tests for PDF invoice header span geometry."""

from __future__ import annotations

import pytest

pytestmark = pytest.mark.unit

from docwen_plugin_optimizer_invoice_cn.invoice_cn.rows import (
    parse_invoice_rows_from_pdf_spans,
)


def _span(x0: float, y0: float, x1: float, text: str) -> tuple[float, float, float, float, str]:
    return (x0, y0, x1, y0 + 8.0, text)


def test_combined_amount_and_tax_rate_header_preserves_right_side_columns() -> None:
    spans = [
        _span(50.0, 100.0, 90.0, "项目名称"),
        _span(150.0, 100.0, 190.0, "规格型号"),
        _span(230.0, 100.0, 250.0, "单位"),
        _span(300.0, 100.0, 325.0, "数 量"),
        _span(370.0, 100.0, 395.0, "单 价"),
        _span(450.0, 100.0, 530.0, "金  额   税率/征收率 "),
        _span(568.0, 100.0, 591.0, "税  额"),
        _span(55.0, 112.0, 95.0, "*服务甲"),
        _span(235.0, 112.0, 250.0, "项"),
        _span(310.0, 112.0, 325.0, "2"),
        _span(379.0, 112.0, 405.0, "10.00"),
        _span(453.0, 112.0, 480.0, "20.00"),
        _span(492.0, 112.0, 510.0, "3%"),
        _span(576.0, 112.0, 600.0, "0.60"),
        _span(55.0, 140.0, 90.0, "价税合计"),
    ]

    rows = parse_invoice_rows_from_pdf_spans(spans=spans)

    assert len(rows) == 1
    assert rows[0]["单价"] == "10.00"
    assert rows[0]["金额"] == "20.00"
    assert rows[0]["税率"] == "3%"
    assert rows[0]["税额"] == "0.60"


def test_character_split_headers_reconstruct_all_column_boundaries() -> None:
    spans = [
        _span(45.0, 100.0, 81.0, "项目名称"),
        _span(119.0, 100.0, 155.0, "规格型号"),
        _span(190.0, 100.0, 199.0, "单"),
        _span(208.0, 100.0, 217.0, "位"),
        _span(264.0, 100.0, 273.0, "数"),
        _span(282.0, 100.0, 291.0, "量"),
        _span(335.0, 100.0, 344.0, "单"),
        _span(353.0, 100.0, 362.0, "价"),
        _span(407.0, 100.0, 416.0, "金"),
        _span(425.0, 100.0, 434.0, "额"),
        _span(446.0, 100.0, 496.0, "税率征收率"),
        _span(464.0, 100.0, 469.0, "/"),
        _span(551.0, 100.0, 560.0, "税"),
        _span(569.0, 100.0, 578.0, "额"),
        _span(50.0, 112.0, 95.0, "*项目乙"),
        _span(195.0, 112.0, 210.0, "件"),
        _span(270.0, 112.0, 280.0, "1"),
        _span(340.0, 112.0, 370.0, "40.00"),
        _span(412.0, 112.0, 442.0, "40.00"),
        _span(465.0, 112.0, 485.0, "6%"),
        _span(558.0, 112.0, 585.0, "2.40"),
        _span(45.0, 140.0, 85.0, "价税合计"),
    ]

    rows = parse_invoice_rows_from_pdf_spans(spans=spans)

    assert len(rows) == 1
    assert rows[0]["单位"] == "件"
    assert rows[0]["数量"] == "1"
    assert rows[0]["单价"] == "40.00"
    assert rows[0]["金额"] == "40.00"
    assert rows[0]["税率"] == "6%"
    assert rows[0]["税额"] == "2.40"


def test_exact_header_spans_remain_supported() -> None:
    labels = ["项目名称", "规格型号", "单位", "数量", "单价", "金额", "税率", "税额"]
    x_starts = [45.0, 125.0, 205.0, 265.0, 325.0, 385.0, 455.0, 535.0]
    spans = [_span(x, 100.0, x + 35.0, label) for x, label in zip(x_starts, labels, strict=True)]
    spans.extend(
        [
            _span(50.0, 112.0, 90.0, "*服务丙"),
            _span(210.0, 112.0, 225.0, "次"),
            _span(270.0, 112.0, 280.0, "1"),
            _span(330.0, 112.0, 355.0, "8.00"),
            _span(390.0, 112.0, 415.0, "8.00"),
            _span(460.0, 112.0, 475.0, "1%"),
            _span(540.0, 112.0, 565.0, "0.08"),
            _span(45.0, 140.0, 85.0, "价税合计"),
        ]
    )

    rows = parse_invoice_rows_from_pdf_spans(spans=spans)

    assert len(rows) == 1
    assert rows[0]["金额"] == "8.00"
    assert rows[0]["税率"] == "1%"
    assert rows[0]["税额"] == "0.08"
