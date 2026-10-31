"""Current field owners for FA-13; real publication regressions run separately."""

from __future__ import annotations

import ast
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

ROOT = Path(__file__).resolve().parents[2]
INVOICE = "packages/plugins/optimizers/invoice_cn/src/docwen_plugin_optimizer_invoice_cn/invoice_cn"


def _tree(relative_path: str) -> ast.Module:
    return ast.parse((ROOT / relative_path).read_text(encoding="utf-8"))


def _function(tree: ast.AST, name: str) -> ast.FunctionDef:
    [function] = [node for node in ast.walk(tree) if isinstance(node, ast.FunctionDef) and node.name == name]
    return function


def _calls(tree: ast.AST, name: str) -> list[ast.Call]:
    return [
        node
        for node in ast.walk(tree)
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id == name
    ]


def test_fa13_repairs_have_non_private_synthetic_regressions() -> None:
    rows = _tree(f"{INVOICE}/rows.py")
    metadata = _tree(f"{INVOICE}/metadata.py")
    spans = _tree(f"{INVOICE}/span_fields.py")
    _function(rows, "_detect_header_columns")

    # Positional buyer/seller and amount ownership moved out of metadata.py.
    # The entry point must retain the source page and diagnostic sink when it
    # delegates; a missing region must not become membership in both parties.
    positional = _function(metadata, "parse_invoice_metadata_from_pdf_spans")
    for name, arguments in (("party_metadata", ("items",)), ("amount_metadata", ("items", "spans"))):
        [call] = _calls(positional, name)
        assert len(call.args) == len(arguments)
        assert tuple(node.id for node in call.args if isinstance(node, ast.Name)) == arguments
        keywords = {keyword.arg: keyword.value for keyword in call.keywords}
        assert isinstance(keywords["sink"], ast.Name) and keywords["sink"].id == "diagnostic_sink"
        assert isinstance(keywords["location"], ast.Name) and keywords["location"].id == "location"
        _function(spans, name)
    parties = _function(spans, "party_metadata")
    assert _calls(parties, "_party_regions")
    assert _calls(parties, "report_metadata_issue")
    assert {
        "INVOICE-PARTY-AMBIGUOUS",
        "INVOICE-METADATA-CONFLICT",
    } <= {node.value for node in ast.walk(spans) if isinstance(node, ast.Constant) and isinstance(node.value, str)}
    amounts = _function(spans, "amount_metadata")
    assert _calls(amounts, "_normalize_ocr_amount")
    assert _calls(amounts, "_detect_header_columns")

    # Compact text and positional fields share signed-token normalization.
    # A regex spelling that truncates thousands or skips a sign is no oracle.
    assert any(
        isinstance(node, ast.ImportFrom)
        and node.module == "docwen_plugin_optimizer_invoice_cn.invoice_cn.ocr_normalize"
        and {"AMOUNT_TOKEN_PATTERN", "_normalize_ocr_amount"} <= {alias.name for alias in node.names}
        for node in metadata.body
    )
    compact = _function(metadata, "parse_invoice_metadata_from_compact_text")
    assert _calls(compact, "_normalize_ocr_amount")
    assert any(
        isinstance(node, ast.JoinedStr)
        and {"AMOUNT_TOKEN_PATTERN", "_MONEY_END"}
        <= {
            value.value.id
            for value in node.values
            if isinstance(value, ast.FormattedValue) and isinstance(value.value, ast.Name)
        }
        for node in ast.walk(compact)
    )

    regressions = {
        "packages/plugins/optimizers/invoice_cn/tests/test_invoice_pdf_span_rows.py": (
            "test_combined_amount_and_tax_rate_header_preserves_right_side_columns",
            "test_character_split_headers_reconstruct_all_column_boundaries",
        ),
        "packages/plugins/optimizers/invoice_cn/tests/test_invoice_pdf_span_metadata.py": (
            "test_currency_prefixed_grand_total_is_selected_from_total_line",
            "test_split_tax_id_labels_infer_distinct_buyer_and_seller_columns",
        ),
        "tests/integration/test_invoice_native_field_identity.py": (
            "test_signed_amount_fields_follow_roles_not_magnitude_or_positive_unit_price",
            "test_party_ids_have_one_proven_role_in_the_published_yaml",
            "test_same_coordinates_on_different_pages_never_share_party_or_detail_rows",
            "test_conflicting_page_metadata_keeps_first_owned_value_and_reports_both_pages",
            "test_complete_party_text_is_not_replaced_by_conflicting_geometry",
        ),
        "tests/integration/test_invoice_complete_amount_tokens.py": (
            "test_complete_amount_token_survives_pdf_runtime_publication",
            "test_invalid_amount_is_not_a_valid_numeric_prefix",
        ),
    }
    for path, names in regressions.items():
        tree = _tree(path)
        for name in names:
            _function(tree, name)
