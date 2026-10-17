"""The line-per-paragraph setting is a typed, persistent GUI choice."""

import pytest

from docwen_gui.models.settings_config import FormattingConfig

pytestmark = pytest.mark.unit


def test_new_setting_defaults_off() -> None:
    config = FormattingConfig()
    assert config.each_line_paragraph is False


def test_new_setting_is_independent_of_dialect_extension_flags() -> None:
    config = FormattingConfig(each_line_paragraph=True)
    assert config.each_line_paragraph is True
    assert not any(config.markdown_extensions["input"].values())
