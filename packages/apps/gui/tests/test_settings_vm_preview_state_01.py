"""Focused tests split from test_settings_vm_preview_state.py."""

from __future__ import annotations

import pytest

from ._settings_vm_preview_state_support import (
    SECTION_GUI,
    ConfigPortAdapter,
    SettingsViewModel,
)
from ._settings_vm_preview_state_support import (
    config_port as config_port,
)

pytestmark = pytest.mark.unit
from ._settings_vm_preview_state_support import (
    vm as vm,
)


class TestBaselineCapture:
    """Opening settings captures the persisted snapshot as baseline."""

    def test_begin_session_captures_persisted_snapshot(self, vm: SettingsViewModel) -> None:
        vm.set_field(SECTION_GUI, "theme", "dark")
        vm.apply_settings()
        # After Apply, persisted == draft == "dark"
        assert vm.config.gui.theme == "dark"

        vm.begin_session()
        # After begin_session, the baseline should be locked at "dark"
        # Editing draft should NOT affect the locked baseline
        vm.set_field(SECTION_GUI, "theme", "light")
        assert vm.config.gui.theme == "light"  # draft changed

        vm.cancel_changes()
        # Cancel should restore to the locked baseline, which was "dark"
        assert vm.config.gui.theme == "dark"

    def test_begin_session_locks_baseline_even_without_controller(self) -> None:
        vm = SettingsViewModel(controller=None)
        vm.begin_session()
        assert vm.config.gui.theme == "light"
        vm.set_field(SECTION_GUI, "theme", "dark")
        vm.cancel_changes()
        assert vm.config.gui.theme == "light"


class TestDraftEditing:
    """Editing must update draft without persisting to config on disk."""

    def test_editing_theme_does_not_persist(self, vm: SettingsViewModel, config_port: ConfigPortAdapter) -> None:
        vm.apply_settings()  # persist initial baseline
        vm.begin_session()
        vm.set_field(SECTION_GUI, "theme", "dark")

        # Check config on disk is NOT yet changed
        theme_on_disk = config_port.get("gui.theme.default_theme")
        assert theme_on_disk != "dark", f"Expected theme NOT to be 'dark' on disk, got {theme_on_disk}"

    def test_apply_persists_after_draft_edits(self, vm: SettingsViewModel, config_port: ConfigPortAdapter) -> None:
        vm.begin_session()
        vm.set_field(SECTION_GUI, "theme", "dark")
        vm.apply_changes()

        theme_on_disk = config_port.get("gui.theme.default_theme")
        assert theme_on_disk == "dark"

    def test_cancel_discards_draft(self, vm: SettingsViewModel) -> None:
        vm.set_field(SECTION_GUI, "theme", "dark")
        vm.apply_settings()
        vm.begin_session()
        vm.set_field(SECTION_GUI, "theme", "light")
        vm.cancel_changes()
        assert vm.config.gui.theme == "dark"


class TestPreviewSemantics:
    """Preview must apply visual changes without persisting or dirtying draft."""

    def test_preview_theme_does_not_change_dirty(self, vm: SettingsViewModel) -> None:
        vm.begin_session()
        vm.preview_theme("dark")
        # Preview-only changes should NOT count as draft edits
        # (theme preview is applied to UI but reverts on Cancel)
        assert vm.config.gui.theme == "light"  # draft still at baseline

    def test_preview_theme_does_not_persist(self, vm: SettingsViewModel, config_port: ConfigPortAdapter) -> None:
        vm.apply_settings()  # persist "light"
        vm.begin_session()
        vm.preview_theme("dark")
        on_disk = config_port.get("gui.theme.default_theme")
        assert on_disk != "dark", f"Preview must not persist, but disk has {on_disk}"

    def test_cancel_restores_preview_to_persisted(self, vm: SettingsViewModel) -> None:
        vm.set_field(SECTION_GUI, "theme", "dark")
        vm.apply_settings()
        vm.begin_session()
        vm.preview_theme("light")
        vm.cancel_changes()
        # After cancel, theme should be back to persisted baseline = "dark"
        assert vm.config.gui.theme == "dark"


class TestApplyUpdatesBaseline:
    """Apply must update the locked baseline so Cancel-after-Apply restores correctly."""

    def test_apply_then_edit_then_cancel_restores_to_post_apply(self, vm: SettingsViewModel) -> None:
        vm.set_field(SECTION_GUI, "theme", "dark")
        vm.apply_settings()
        vm.begin_session()
        vm.set_field(SECTION_GUI, "theme", "light")
        vm.apply_changes()
        # Now baseline should be "light" (post-Apply)
        vm.set_field(SECTION_GUI, "theme", "system")
        vm.cancel_changes()
        # Cancel should restore to post-Apply baseline = "light", NOT the
        # older "dark" from before begin_session
        assert vm.config.gui.theme == "light"

    def test_ok_applies_then_signals_close_readiness(
        self, vm: SettingsViewModel, config_port: ConfigPortAdapter
    ) -> None:
        vm.begin_session()
        vm.set_field(SECTION_GUI, "theme", "dark")
        result = vm.ok_changes()
        assert result is True
        assert vm.is_dirty is False
        assert config_port.get("gui.theme.default_theme") == "dark"


class TestOpacityPreview:
    """Opacity preview must apply visual effect without persisting."""

    def test_preview_opacity_does_not_persist(self, vm: SettingsViewModel, config_port: ConfigPortAdapter) -> None:
        vm.apply_settings()
        vm.begin_session()
        vm.preview_opacity(0.5)
        on_disk = config_port.get("gui.transparency.default_value")
        assert on_disk != 0.5, f"Opacity preview must not persist, but disk has {on_disk}"

    def test_cancel_restores_opacity_to_persisted(self, vm: SettingsViewModel) -> None:
        vm.set_field(SECTION_GUI, "transparency_value", 0.8)
        vm.apply_settings()
        vm.begin_session()
        vm.preview_opacity(0.3)
        vm.cancel_changes()
        assert vm.config.gui.transparency_value == 0.8


class TestCloseWithoutApply:
    """Closing the dialog without Apply must not modify persisted config."""

    def test_close_discards_all_drafts(self, vm: SettingsViewModel) -> None:
        vm.set_field(SECTION_GUI, "theme", "dark")
        vm.apply_settings()
        vm.begin_session()
        vm.set_field(SECTION_GUI, "theme", "light")
        vm.set_field(SECTION_GUI, "language", "en_US")
        vm.cancel_changes()
        # All draft changes discarded
        assert vm.config.gui.theme == "dark"
        assert vm.config.gui.language == "zh_CN"


class TestDirtyState:
    """Dirty tracking must reflect draft vs persisted, not preview."""

    def test_preview_only_does_not_make_dirty(self, vm: SettingsViewModel) -> None:
        vm.begin_session()
        vm.preview_theme("dark")
        assert vm.is_dirty is False

    def test_edit_makes_dirty(self, vm: SettingsViewModel) -> None:
        vm.begin_session()
        vm.set_field(SECTION_GUI, "theme", "dark")
        assert vm.is_dirty is True

    def test_apply_clears_dirty(self, vm: SettingsViewModel) -> None:
        vm.begin_session()
        vm.set_field(SECTION_GUI, "theme", "dark")
        assert vm.is_dirty is True
        vm.apply_changes()
        assert vm.is_dirty is False

    def test_cancel_clears_dirty(self, vm: SettingsViewModel) -> None:
        vm.begin_session()
        vm.set_field(SECTION_GUI, "theme", "dark")
        assert vm.is_dirty is True
        vm.cancel_changes()
        assert vm.is_dirty is False
