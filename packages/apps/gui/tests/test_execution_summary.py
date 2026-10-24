"""The user sees normalized request values, never guessed defaults or secrets."""

from __future__ import annotations

import pytest

from docwen_core.models.request import POSTPROCESS_PROOFREAD_OPTION, ConversionRequest, OutputPolicy
from docwen_gui.execution_summary import request_summary
from docwen_gui.i18n import get_locale, set_locale

pytestmark = pytest.mark.unit


def test_summary_names_only_effective_requested_options():
    previous = get_locale()
    set_locale("en_US")
    try:
        request = ConversionRequest(
            request_id="summary",
            input_refs=[],
            target_format="docx",
            options={
                "template_name": "opaque-id",
                "remove_numbering": True,
                "add_numbering": False,
                "numbering_scheme": "unused-scheme",
                "spreadsheet_password": "secret-never-show",
                POSTPROCESS_PROOFREAD_OPTION: {"enable_typos_rule": True, "enable_symbol_pairing": False},
            },
            output_policy=OutputPolicy(output_dir="/exports", date_subfolder="iso"),
        )
        summary = request_summary(request, template_label="Annual report")
        assert "Template: Annual report" in summary
        assert "Remove heading numbers: On" in summary
        assert "Add heading numbers: Off" in summary
        assert "Typo Correction" in summary
        assert "Symbol Pairing" not in summary
        assert "/exports" in summary
        assert "dated subfolder" in summary
        assert all(value not in summary for value in ("opaque-id", "unused-scheme", "secret-never-show"))
        request.options = {}
        assert "heading numbers" not in request_summary(request)
    finally:
        set_locale(previous)
