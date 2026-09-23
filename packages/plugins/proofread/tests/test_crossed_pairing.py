"""Crossed nesting must not be mistaken for independent balanced symbol types."""

from __future__ import annotations

import pytest

from docwen_plugin_proofread.text_validator import TextValidator

pytestmark = pytest.mark.unit


@pytest.mark.parametrize("text,position", [("([)]", 2), ("（【）】", 2), ("{(})", 2)])
def test_crossed_nesting_reports_the_closing_symbol(text, position):
    errors = TextValidator().validate_text(text)
    paired = [error for error in errors if error.source == "pairing"]
    assert len(paired) == 1
    assert paired[0].start_pos == position
    assert paired[0].end_pos == position + 1
    assert paired[0].error_text == text[position]
    assert paired[0].replacement is None


@pytest.mark.parametrize("text", ["([])", "（【正文】）", "{[()]}", "don't", "dogs' owners", "‘don't’", '"([text])"'])
def test_valid_nesting_and_apostrophes_are_not_errors(text):
    assert not [error for error in TextValidator().validate_text(text) if error.source == "pairing"]
