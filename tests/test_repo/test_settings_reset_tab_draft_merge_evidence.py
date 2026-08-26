"""Guards for VIS-2026-07-19-138 Settings Reset Tab draft reconciliation."""

from __future__ import annotations

from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "settings-reset-tab-draft-merge-2026-07-19.md"


def _read(relative_path: str) -> str:
    return read_source_text(PROJECT_ROOT / relative_path)


def _method_block(source: str, name: str, next_name: str) -> str:
    start = source.index(f"    def {name}(")
    end = source.index(f"    def {next_name}(", start)
    return source[start:end]


def test_reset_tab_draft_merge_stays_model_owned_and_preserves_unrelated_drafts() -> None:
    vm = _read("packages/apps/gui/src/docwen_gui/view_models/settings_vm.py")
    text_tab = _read("packages/apps/gui/src/docwen_gui/widgets/settings/text_tab.py")
    tabbed = _read("packages/apps/gui/src/docwen_gui/widgets/template_selector_tabbed.py")
    vm_regression = _read("packages/apps/gui/tests/test_settings_vm_preview_state_*.py")
    dialog_regression = _read("packages/apps/gui/tests/test_settings_dialog_shell_*.py")

    ownership = vm.split(
        "_RESET_GROUP_DRAFT_PATHS: dict[str, tuple[tuple[str, ...], ...]] = {",
        1,
    )[1].split("_MISSING = object()", 1)[0]
    for group in (
        "general",
        "text",
        "proofread",
        "document",
        "spreadsheet",
        "image",
        "layout",
        "link",
        "formatting",
        "output",
        "export",
        "logging",
        "other",
        "conversion_defaults",
        "software_priority",
        "software",
    ):
        assert f'"{group}": (' in ownership

    assert '("gui", "md_default_template")' not in ownership.split('"text": (', 1)[0]
    assert '("gui", "md_default_template")' in ownership.split('"text": (', 1)[1]
    proofread_ownership = ownership.split('"proofread": (', 1)[1].split('"document": (', 1)[0]
    assert '("proofread", "symbol_mappings")' not in proofread_ownership
    assert '("conversion_defaults", "document", "to_md_keep_images")' in ownership
    assert '("software_priority", "word_processors")' in ownership
    assert '("conversion_defaults", "export")' in ownership
    assert '("formatting", "table_style_mode")' in ownership
    assert '"software": (("software_priority",),)' in ownership
    assert "docwen_runtime" not in ownership
    assert "reset_plan_for_group" not in ownership

    reconcile = _method_block(vm, "_reconcile_reset_attempt", "_merge_group_reset_from_raw")
    merge = _method_block(vm, "_merge_group_reset_from_raw", "_mark_current_config_as_persisted_baseline")
    reset_group = _method_block(vm, "reset_group", "reset_all")
    reset_all = _method_block(vm, "reset_all", "_persist_to_controller_config")

    assert "source_may_have_changed = before_source is None or after_source != before_source" in reconcile
    assert "group is None or (operation_succeeded and group not in _RESET_GROUP_DRAFT_PATHS)" in reconcile
    assert "self._merge_group_reset_from_raw(" in reconcile
    assert "source_paths = _changed_model_paths(before, after)" in merge
    assert "reset_paths = set(source_paths)" in merge
    assert "reset_paths.update(_RESET_GROUP_DRAFT_PATHS.get(group, ()))" in merge
    assert "draft_paths = _changed_model_paths(previous_snapshot, previous_draft)" in merge
    assert "if any(_model_paths_overlap(path, reset_path) for reset_path in reset_paths):" in merge
    assert "_write_model_path(merged, path, _read_model_path(previous_draft, path))" in merge
    assert "_mark_model_dirty_against(merged, after)" in merge
    assert "self._snapshot = deepcopy(after)" in merge
    assert "self._persisted_baseline = deepcopy(after)" in merge
    assert "self.config_reloaded.emit()" in merge
    assert 'getattr(controller, "config_port", None)' in reset_group
    assert "operation_succeeded=ok" in reset_group and "group=group" in reset_group
    assert "group=None" in reset_all
    assert "docwen_runtime" not in reset_group

    for token in (
        "test_failed_reset_preserves_draft_or_reloads_partial_source",
        "test_successful_group_reset_preserves_non_owner_drafts",
        "test_successful_proofread_reset_preserves_protected_dictionary_draft",
        "test_successful_document_reset_preserves_sibling_conversion_drafts",
        "test_reset_tab_draft_ownership_covers_every_dialog_group",
        "test_successful_builtin_group_reset_discards_only_owner_draft",
        "test_successful_unknown_group_uses_conservative_full_refresh",
        "test_conversion_defaults_aggregate_resets_owned_text_and_export_leaves",
        "test_software_alias_resets_software_priority_draft",
    ):
        assert token in vm_regression

    for token in (
        "test_settings_dialog_reset_tab_preserves_other_tab_draft_and_widget",
        "test_settings_text_template_restore_does_not_create_unsaved_draft",
        "test_settings_text_reset_restores_template_type_widget_without_dirty",
    ):
        assert token in dialog_regression

    assert "peek_callback_selection_feedback()" in text_tab
    assert "restore_current_tab(config.gui.md_default_template)" in text_tab
    assert "restore_current_tab(self._vm.config.gui.md_default_template)" in text_tab
    assert "def peek_callback_selection_feedback(" in tabbed
    assert "def restore_current_tab(" in tabbed
    assert "self._selection_callback_contexts.append(callback_feedback)" in tabbed
    assert "self.template_selected.emit(template_type, template_name)" in tabbed
    assert "self._selection_callback_contexts.pop()" in tabbed
