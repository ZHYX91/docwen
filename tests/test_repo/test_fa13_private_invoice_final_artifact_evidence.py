"""Fail-closed evidence guards for VIS-386 private invoice parity."""

from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "fa13-private-invoice-final-artifact-parity-2026-07-26.md"
CARD_NAME = "fa13-private-invoice-corpus-stage-card-2026-07-26.md"
ADDENDUM_NAME = "fa13-private-invoice-execution-addendum-2026-07-26.md"
STATUS = "FIXED_AND_VERIFIED_WITH_USER_ACCEPTED_DERIVED_IMAGE_ACCURACY_BOUNDARY"


def _read(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def test_fa13_repairs_have_non_private_synthetic_regressions() -> None:
    rows_source = _read(
        ROOT
        / "packages"
        / "plugins"
        / "optimizers"
        / "invoice_cn"
        / "src"
        / "docwen_plugin_optimizer_invoice_cn"
        / "invoice_cn"
        / "rows.py"
    )
    metadata_source = _read(
        ROOT
        / "packages"
        / "plugins"
        / "optimizers"
        / "invoice_cn"
        / "src"
        / "docwen_plugin_optimizer_invoice_cn"
        / "invoice_cn"
        / "metadata.py"
    )
    row_tests = _read(
        ROOT / "packages" / "plugins" / "optimizers" / "invoice_cn" / "tests" / "test_invoice_pdf_span_rows.py"
    )
    metadata_tests = _read(
        ROOT / "packages" / "plugins" / "optimizers" / "invoice_cn" / "tests" / "test_invoice_pdf_span_metadata.py"
    )

    assert "_detect_header_columns" in rows_source
    assert "character positions" in rows_source
    assert "largest_gap >= 50.0" in metadata_source
    assert 'r"[¥￥]?-?[0-9]+(?:\\.[0-9]{1,2})?"' in metadata_source
    assert "test_combined_amount_and_tax_rate_header_preserves_right_side_columns" in row_tests
    assert "test_character_split_headers_reconstruct_all_column_boundaries" in row_tests
    assert "test_currency_prefixed_grand_total_is_selected_from_total_line" in metadata_tests
    assert "test_split_tax_id_labels_infer_distinct_buyer_and_seller_columns" in metadata_tests
