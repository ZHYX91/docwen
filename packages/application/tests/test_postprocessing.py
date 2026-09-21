"""Post-processing ownership and non-regression contracts."""

from __future__ import annotations

import pytest

from docwen_application.postprocessing import prepare_postprocess_options

pytestmark = pytest.mark.unit


def test_named_markdown_validate_keeps_route_owned_proofread_options() -> None:
    options = {
        "enable_symbol_pairing": True,
        "enable_symbol_correction": False,
        "enable_typos_rule": True,
        "enable_sensitive_word": False,
    }

    prepared = prepare_postprocess_options(
        options,
        source_format="markdown",
        target_format="markdown",
        action_name="validate",
    )

    assert prepared == options
    assert prepared is not options


def test_conversion_postprocess_requires_explicit_or_enabled_checks() -> None:
    assert prepare_postprocess_options(
        {
            "enable_symbol_pairing": False,
            "enable_symbol_correction": False,
            "enable_typos_rule": False,
            "enable_sensitive_word": False,
        },
        source_format="markdown",
        target_format="docx",
        action_name="",
    ) == {}
