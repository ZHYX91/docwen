"""Settings change-summary behavior.

The Settings dialog shows field-level old/new changes while edits are still a
draft. These tests keep that UX contract separate from screenshot evidence.
"""

from __future__ import annotations

import pytest

from docwen_gui.view_models.settings_vm import SECTION_GUI, SECTION_OUTPUT, SettingsViewModel
from docwen_gui.widgets.settings.dialog import (
    _abbreviate_value,
    _friendly_change_field,
    _friendly_change_value,
)

pytestmark = pytest.mark.unit


def _changes_by_field(vm: SettingsViewModel) -> dict[str, tuple[object, object]]:
    return {str(change["field"]): (change["old"], change["new"]) for change in vm.get_change_summary()}


def test_change_summary_reports_field_level_old_new_dotted_paths() -> None:
    vm = SettingsViewModel()
    vm.begin_session()

    vm.set_field(SECTION_GUI, "theme", "dark")
    vm.set_field(SECTION_OUTPUT, "output_mode", "custom")

    changes = _changes_by_field(vm)

    assert changes["gui.theme"] == ("light", "dark")
    assert changes["output.output_mode"] == ("source", "custom")
    assert "gui" not in changes
    assert "output" not in changes
    assert all(not field.startswith("_") for field in changes)


def test_apply_clears_change_summary_and_updates_baseline() -> None:
    vm = SettingsViewModel()
    vm.begin_session()
    vm.set_field(SECTION_GUI, "theme", "dark")
    assert _changes_by_field(vm) == {"gui.theme": ("light", "dark")}

    assert vm.apply_changes() is True

    assert vm.get_change_summary() == []
    vm.set_field(SECTION_GUI, "theme", "system")
    assert _changes_by_field(vm) == {"gui.theme": ("dark", "system")}


def test_cancel_clears_change_summary_and_restores_draft() -> None:
    vm = SettingsViewModel()
    vm.begin_session()
    vm.set_field(SECTION_GUI, "theme", "dark")
    assert vm.get_change_summary()

    vm.cancel_changes()

    assert vm.config.gui.theme == "light"
    assert vm.get_change_summary() == []


def test_change_summary_tooltip_values_are_abbreviated() -> None:
    long_value = "x" * 80

    assert _abbreviate_value(long_value) == ("x" * 57) + "..."
    assert _abbreviate_value("short") == "short"


def test_change_summary_projects_known_fields_and_values_but_keeps_honest_fallbacks() -> None:
    from docwen_gui.i18n import t

    assert _friendly_change_field("gui.theme") == t("settings.general.theme_label").rstrip(":：")
    assert _friendly_change_value("gui.theme", "dark") == t("settings.general.themes.dark")
    assert _friendly_change_value("output.create_date_subfolder", True) == t("settings.changes.enabled")
    unknown = _friendly_change_field("future.section.value")
    assert "future.section.value" in unknown
