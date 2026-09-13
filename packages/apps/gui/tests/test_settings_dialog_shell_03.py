"""Template state remains independent of settings drafts."""

import pytest

pytestmark = pytest.mark.gui


def test_template_page_is_part_of_settings_navigation():
    from docwen_gui.widgets.settings.dialog import TAB_KEYS, TAB_NAMES

    assert "templates" in TAB_KEYS
    assert TAB_NAMES["templates"]
