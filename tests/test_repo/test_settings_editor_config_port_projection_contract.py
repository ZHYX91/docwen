"""Fail-closed evidence guards for VIS-2026-07-19-141 editor config ports."""

from __future__ import annotations

import re
from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "settings-editor-config-port-composition-2026-07-19.md"


def _read(relative_path: str) -> str:
    return read_source_text(PROJECT_ROOT / relative_path)


def _method_block(source: str, name: str, next_name: str) -> str:
    start = source.index(f"    def {name}(")
    end = source.index(f"    def {next_name}(", start)
    return source[start:end]


def test_settings_editor_persistence_stays_on_the_injected_config_port() -> None:
    port = _read("packages/application/src/docwen_application/ports/runtime.py")
    adapter = _read("packages/bundle/src/docwen_bundle/config_port.py")
    vm = _read("packages/apps/gui/src/docwen_gui/view_models/settings_vm.py")
    text_tab = _read("packages/apps/gui/src/docwen_gui/widgets/settings/text_tab.py")
    proofread_tab = _read("packages/apps/gui/src/docwen_gui/widgets/settings/proofread_tab.py")
    add_editor = _read("packages/apps/gui/src/docwen_gui/widgets/settings/numbering_add_editor.py")
    clean_editor = _read("packages/apps/gui/src/docwen_gui/widgets/settings/numbering_clean_editor.py")

    assert "def get_file_text(self, rel_path: str) -> str | None:" in port
    assert "def save_file_text(self, rel_path: str, content: str) -> bool:" in port
    assert "return self._loader.get_file_text(rel_path)" in adapter
    assert "return self._loader.save_file_text(rel_path, content)" in adapter

    ownership = vm.split(
        "_EDITOR_FILE_MODEL_PATHS: dict[str, tuple[tuple[str, ...], ...]] = {",
        1,
    )[1].split("_MISSING = object()", 1)[0]
    for rel_path in (
        "numbering/add.toml",
        "numbering/cleanup.toml",
        "proofread/pairs.toml",
        "proofread/symbol_map.toml",
        "proofread/typos.toml",
        "proofread/sensitive_words.toml",
    ):
        assert f'"{rel_path}"' in ownership
    assert '"gui.toml"' not in ownership
    assert "get_config_loader" not in vm

    read_source = _method_block(vm, "read_config_file_text", "save_config_file_text")
    save_source = _method_block(vm, "save_config_file_text", "make_save_config_text_callback")
    reconcile = _method_block(vm, "_reconcile_editor_file_source", "_persist_editor_values")
    assert "config_name not in _EDITOR_FILE_MODEL_PATHS" in read_source
    assert 'getattr(cfg_port, "get_file_text", None)' in read_source
    assert "config_name not in _EDITOR_FILE_MODEL_PATHS" in save_source
    assert 'getattr(cfg_port, "save_file_text", None)' in save_source
    assert "self._reconcile_editor_file_source(config_name, after)" in save_source
    assert "_write_model_path(self._snapshot, path, value)" in reconcile
    assert "_write_model_path(self._persisted_baseline, path, value)" in reconcile

    assert "get_config_loader" not in proofread_tab
    assert "update_file_dict" not in proofread_tab
    assert "update_file_value" not in proofread_tab
    open_editor = _method_block(proofread_tab, "_open_config_editor", "_open_symbol_mapping_editor")
    assert "self._vm.read_config_file_text(config_name)" in open_editor
    assert "save_callback=self._vm.save_config_file_text" in open_editor

    assert text_tab.count("config_data=") >= 4
    assert "config_data=self._vm.config.text.numbering_schemes" in text_tab
    assert "if callable(self._on_save) and self._on_save(doc_data) is False:" in add_editor
    assert "if callable(self._on_save) and self._on_save(doc_data) is False:" in clean_editor


def test_editor_replacement_sections_have_one_marker_free_semantics() -> None:
    registry = _read("packages/runtime/src/docwen_runtime/config/registry.py")
    loader = _read("packages/runtime/src/docwen_runtime/config/loader.py")

    assert "replace_sections: frozenset[str] = frozenset()" in registry
    for rel_path in (
        "proofread/symbol_map.toml",
        "proofread/typos.toml",
        "proofread/sensitive_words.toml",
    ):
        assert re.search(
            rf'ConfigFileSpec\(\s*"{re.escape(rel_path)}".*?'
            r'replace_sections=frozenset\(\{"entries"\}\),\s*\)',
            registry,
            re.DOTALL,
        )
    assert 'ConfigFileSpec("proofread/pairs.toml"' in registry

    for token in (
        'str(key).startswith("__docwen_")',
        "result = deep_merge(deepcopy(base), deepcopy(user))",
        "if section in spec.replace_sections or section not in user:",
        "for section in spec.replace_sections:",
        "spec.replace_sections,",
        "complete_sections = frozenset(file_spec.replace_sections).intersection(",
        "_MISSING_CONFIG_VALUE = object()",
        "base_value = _read_nested_value(base, rel_key)",
        "base_value is _MISSING_CONFIG_VALUE",
        "preserve_top_level=True",
    ):
        assert token in loader
    for removed_token in (
        "user_replace_sections",
        "_USER_REPLACE_SECTIONS_MARKER",
        "_declared_user_replace_sections",
        "_user_layer_without_marker",
        "_mark_complete_user_sections",
        "_replace_complete_user_sections_marker",
        "_sync_user_replace_sections_marker",
    ):
        assert removed_token not in loader
    assert "base_value is missing" not in loader


def test_direct_editor_and_loader_regressions_remain_named_and_discoverable() -> None:
    gui_regression = _read("packages/apps/gui/tests/test_settings_editor_config_port.py")
    loader_regression = _read("packages/runtime/tests/test_config_loader_*.py")
    save_regression = _read("packages/runtime/tests/test_save_file_text.py")

    for test_name in (
        "test_text_numbering_editors_write_the_injected_config_port",
        "test_text_settings_apply_writes_only_the_injected_config_port",
        "test_custom_numbering_scheme_deletion_round_trips",
        "test_proofread_editor_reads_effective_base_data_through_the_injected_port",
        "test_proofread_editor_writes_the_injected_port_not_an_unrelated_loader",
        "test_proofread_editor_save_survives_later_settings_apply",
        "test_immediate_proofread_save_updates_cancel_baseline",
        "test_proofread_editor_can_delete_a_shipped_dictionary_entry",
        "test_editor_save_failures_keep_child_dialogs_open",
        "test_settings_vm_rejects_unowned_editor_file_before_writing",
    ):
        assert f"def {test_name}(" in gui_regression
    assert "unrelated_loader = ConfigLoader(" in gui_regression
    assert "assert unrelated_loader.config.as_dict()" in gui_regression
    assert "get_config_loader" not in gui_regression

    for test_name in (
        "test_missing_user_file_returns_shipped_source_verbatim",
        "test_sparse_user_source_overlays_base_with_comments",
        "test_curated_entries_replace_base_so_editor_deletion_round_trips",
        "test_existing_curated_user_section_is_complete_across_shipped_upgrades",
        "test_empty_curated_section_suppresses_shipped_entries",
        "test_replacement_wins_if_a_future_section_is_also_keyed",
        "test_whole_section_writes_use_registry_replacement",
        "test_batched_leaf_writes_create_one_complete_replacement",
        "test_reset_complete_section_reveals_base_then_leaf_write_starts_new_replacement",
        "test_reset_replacement_base_leaf_materializes_default_without_reviving_siblings",
        "test_reset_replacement_custom_leaf_keeps_shipped_entries_deleted",
        "test_leaf_write_creates_complete_replacement_without_hidden_metadata",
        "test_editor_save_writes_no_internal_metadata",
        "test_same_base_and_user_directory_writes_plain_toml",
        "test_removed_internal_marker_is_rejected_fail_closed",
    ):
        assert f"def {test_name}(" in loader_regression
    assert "def test_directory_creation_failure_returns_false(" in save_regression
