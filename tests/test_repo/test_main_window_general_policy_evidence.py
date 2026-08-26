"""Guards for VIS-2026-07-18-137 General MainWindow policy evidence."""

from __future__ import annotations

from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "main-window-general-policy-consumption-2026-07-18.md"


def _read(relative_path: str) -> str:
    return read_source_text(PROJECT_ROOT / relative_path)


def test_main_window_general_policy_source_and_lifecycle_stay_wired() -> None:
    policy = _read("packages/apps/gui/src/docwen_gui/window_behavior.py")
    main = _read("packages/apps/gui/src/docwen_gui/main_window.py")
    dialog = _read("packages/apps/gui/src/docwen_gui/widgets/settings/dialog.py")
    model = _read("packages/apps/gui/src/docwen_gui/models/settings_config.py")
    settings_vm = _read("packages/apps/gui/src/docwen_gui/view_models/settings_vm.py")
    loader = _read("packages/runtime/src/docwen_runtime/config/loader.py")
    regression = _read("packages/apps/gui/tests/test_main_window_window_behavior_*.py")
    dialog_regression = _read("packages/apps/gui/tests/test_settings_dialog_shell_*.py")
    vm_regression = _read("packages/apps/gui/tests/test_settings_vm_preview_state_*.py")
    loader_regression = _read("packages/runtime/tests/test_config_loader_*.py")

    for key in ("remember_gui_state", "auto_center", "expand_side_panels"):
        assert f'"gui.window.{key}"' in policy
        assert f"DEFAULT_WINDOW_BEHAVIOR.{key}" in model
        assert key in main
        assert key in regression

    for token in (
        "@dataclass(frozen=True, slots=True)",
        "return value if isinstance(value, bool) else default",
        "remember_gui_state: bool = True",
        "auto_center: bool = False",
        "expand_side_panels: bool = False",
    ):
        assert token in policy

    assert "if not self._window_behavior.remember_gui_state:" in main
    assert "self._window_behavior.auto_center or not self._window_behavior.remember_gui_state" in main
    assert "2 if expand and left_visible else 0" in main
    assert "3 if expand and right_visible else 0" in main
    assert "cfg_port.set_many(" in main
    assert "settings_source_changed = Signal()" in dialog
    assert "settings_source_changed.connect(self._apply_runtime_window_settings)" in main
    assert "finally:\n            self.settings_source_changed.emit()" in dialog
    assert "gui = self._vm.persisted_config.gui" in dialog
    assert "self._commit_preview_state()" in dialog
    ok_path = dialog.split("def _on_ok", 1)[1].split("def _on_apply", 1)[0]
    apply_path = dialog.split("def _on_apply", 1)[1].split("def _commit_preview_state", 1)[0]
    assert "applied = self._vm.ok_changes()" in ok_path
    assert "if applied:" in ok_path
    assert "self._close_cleanup_done = True" in ok_path
    assert ok_path.index("self._close_cleanup_done = True") < ok_path.index("self.accept()")
    assert "self._vm.apply_changes()" in apply_path
    for handler in (ok_path, apply_path):
        assert "finally:" in handler
        assert "self._commit_preview_state()" in handler
        assert "self.settings_source_changed.emit()" in handler
    assert "self._reload_tab_after_reset(tab_key, tab)" in dialog
    assert 'logger.exception("Unexpected Settings tab reload failure: %s", tab_key)' in dialog
    persist_to_controller = settings_vm.split("def _persist_to_controller_config", 1)[1].split(
        "def _collect_conversion_defaults", 1
    )[0]
    assert "self._collect_markdown_numbering_values(" in persist_to_controller
    assert "self._collect_numbering_clean_values(" in persist_to_controller
    assert persist_to_controller.count("return self._persist_config_values(cfg_port, values)") == 1
    assert "self._refresh_persisted_baseline_from_controller()" in settings_vm
    assert "source_may_have_changed = before_source is None or after_source != before_source" in settings_vm
    transaction_path = loader.split("def _run_user_file_transaction", 1)[1].split("def _mutate_user_document", 1)[0]
    for token in (
        "_CONFIG_TRANSACTION_STATE.nested_attempted = True",
        "preimages = {path: _capture_user_file_preimage(path) for path in paths}",
        "for preimage in reversed(preimages.values()):",
        "_restore_user_file_preimage(preimage, operation=operation)",
        "self.reload()",
    ):
        assert token in transaction_path
    assert "self._config_state_trusted = False" in loader
    assert "self._config_state_trusted = True" in loader
    assert "with contextlib.suppress(Exception)" not in loader

    update_document_path = loader.split("def update_file_document", 1)[1].split("def get_file_dict", 1)[0]
    assert "self._mutate_user_document(spec, mutate)" in update_document_path
    assert "self._run_user_file_transaction(" in update_document_path
    assert 'operation=f"update_file_document:{rel_path}"' in update_document_path

    for token in (
        "test_startup_center_policy_retains_size_only_when_state_is_remembered",
        "test_state_save_reloads_remember_policy_before_writing",
        "test_expand_policy_never_allocates_hidden_columns",
        "test_runtime_policy_refreshes_flags_and_layout_without_recentering",
        "test_settings_dialog_source_signal_refreshes_main_window_policy",
        "ConfigPortAdapter",
    ):
        assert token in regression

    for token in (
        "test_settings_dialog_apply_refreshes_source_after_success_or_partial_failure",
        "test_settings_dialog_partial_apply_uses_persisted_visual_cancel_baseline",
        "test_settings_dialog_failed_noop_reset_keeps_persisted_visual_cancel_baseline",
        "test_settings_dialog_reset_general_updates_visual_cancel_baseline",
        "test_settings_dialog_reset_attempts_refresh_possible_partial_source",
        "test_settings_dialog_reset_reload_failures_are_contained_and_other_tabs_continue",
    ):
        assert token in dialog_regression

    assert "test_failed_reset_preserves_draft_or_reloads_partial_source" in vm_regression
    assert "test_update_file_document_contains_parent_directory_creation_failure" in loader_regression
