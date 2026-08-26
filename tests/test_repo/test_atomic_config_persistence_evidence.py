"""Fail-closed evidence guards for VIS-2026-07-19-145 config integrity."""

from __future__ import annotations

import ast
from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "atomic-config-persistence-and-rollback-2026-07-19.md"


def _read(relative_path: str) -> str:
    return read_source_text(PROJECT_ROOT / relative_path)


def _method_source(relative_path: str, class_name: str, method_name: str) -> str:
    source = _read(relative_path)
    tree = ast.parse(source)
    for node in tree.body:
        if isinstance(node, ast.ClassDef) and node.name == class_name:
            for member in node.body:
                if isinstance(member, (ast.FunctionDef, ast.AsyncFunctionDef)) and member.name == method_name:
                    return ast.get_source_segment(source, member) or ""
    raise AssertionError(f"missing {class_name}.{method_name} in {relative_path}")


def test_runtime_toml_writers_use_same_directory_staged_replace() -> None:
    source = _read("packages/runtime/src/docwen_runtime/toml_io.py")

    for token in (
        "def atomic_write_bytes(",
        "if file_path.is_symlink():",
        "file_path.resolve(strict=False)",
        "tempfile.mkstemp(",
        "dir=file_path.parent",
        'with os.fdopen(descriptor, "wb")',
        "stream.flush()",
        "os.fsync(stream.fileno())",
        "os.replace(temp_path, file_path)",
        "temp_path.unlink()",
        "atomic_write_text(path, tomlkit.dumps(data))",
        "atomic_write_text(path, doc.as_string())",
    ):
        assert token in source
    assert "file_path.write_text(" not in source


def test_every_public_loader_mutation_uses_the_shared_executor() -> None:
    relative_path = "packages/runtime/src/docwen_runtime/config/loader.py"
    source = _read(relative_path)
    transaction = _read("packages/runtime/src/docwen_runtime/config/transaction.py")

    for token in (
        "_CONFIG_TRANSACTION_LOCK = threading.RLock()",
        "_CONFIG_TRANSACTION_STATE = threading.local()",
        "config_transaction.process_config_lock(self._user_dir)",
        "config_transaction.recover_transaction_journal(",
        "config_transaction.write_transaction_journal(",
        "config_transaction.mark_transaction_committed(",
        "_capture_user_file_preimage(path)",
        "Nested configuration persistence rejected",
        "preimages =",
        "for preimage in reversed(preimages.values()):",
        "_restore_user_file_preimage(preimage, operation=operation)",
        "Configuration transaction rollback failed",
        "self._config_state_trusted = False",
        "self._config_state_trusted = True",
        "self._reset_grouped_values_on_disk(dotted_by_file)",
    ):
        assert token in source

    for token in (
        "class UserFilePreimage:",
        "symlink_target: Path | None = None",
        "resolved_target: Path | None = None",
        'CONFIG_JOURNAL_NAME = ".docwen-config.transaction.json"',
        'CONFIG_LOCK_NAME = ".docwen-config.lock"',
        "if state not in {_PREPARED, _COMMITTED}:",
        "if record.state == _COMMITTED:",
        "for preimage in reversed(record.preimages):",
        "Configuration rollback comparison failed; forcing restore",
        "configuration transaction journal checksum mismatch",
        "configuration transaction journal path escapes user directory",
    ):
        assert token in transaction

    for method_name in (
        "reset_file",
        "reset_values",
        "reset_group",
        "reset_all",
        "set_values",
        "write_file_content",
        "save_file_text",
        "update_file_sections",
        "update_file_document",
    ):
        method = _method_source(relative_path, "ConfigLoader", method_name)
        assert "_run_user_file_transaction(" in method, method_name

    assert "return self.set_values({dotted_key: value})" in _method_source(relative_path, "ConfigLoader", "set_value")
    assert "return self.reset_file(spec.rel_path)" in _method_source(relative_path, "ConfigLoader", "reset_section")
    assert "self.reset_file(" not in _method_source(relative_path, "ConfigLoader", "reset_group")
    assert "self.reset_file(" not in _method_source(relative_path, "ConfigLoader", "reset_all")


def test_failure_injection_and_reversed_partial_contracts_remain_discoverable() -> None:
    safety_tests = _read("packages/runtime/tests/test_config_transaction_safety.py")
    loader_tests = _read("packages/runtime/tests/test_config_loader_*.py")

    for test_name in (
        "test_atomic_toml_fsync_failure_preserves_original_and_cleans_temp",
        "test_atomic_toml_replace_failure_keeps_absent_target_absent",
        "test_atomic_toml_write_preserves_symlink_path_and_replaces_resolved_target",
        "test_broken_symlink_write_reload_failure_restores_link_and_missing_target",
        "test_valid_symlink_reset_reload_failure_restores_link_and_target",
        "test_reset_file_removes_broken_symlink_instead_of_leaving_latent_override",
        "test_set_values_second_file_failure_restores_all_preimages",
        "test_reset_all_middle_unlink_failure_restores_all_preimages",
        "test_late_runtime_wiring_failure_restores_disk_and_effective_cache",
        "test_double_reconciliation_failure_marks_effective_state_untrusted",
        "test_nested_persistence_from_document_mutator_fails_outer_operation_closed",
        "test_set_values_planning_failure_returns_false_without_mutation",
        "test_rollback_failure_reports_and_reconciles_actual_disk_state",
        "test_rollback_comparison_read_failure_forces_preimage_restore",
    ):
        assert f"def {test_name}(" in safety_tests

    for stale_name in (
        "test_reset_group_continues_after_one_file_raises",
        "test_reset_all_continues_after_one_owner_raises_and_reconciles_cache",
        "test_set_values_reloads_partial_disk_state_after_later_file_failure",
        "test_set_values_retries_cache_reload_after_transient_reload_failure",
        "test_transient_reload_failure_returns_false_but_reconciles_written_document",
    ):
        assert stale_name not in loader_tests


def test_application_and_gui_callsites_fail_closed_without_sequential_writes() -> None:
    port = _read("packages/application/src/docwen_application/ports/runtime.py")
    settings_vm = _read("packages/apps/gui/src/docwen_gui/view_models/settings_vm.py")
    app = _read("packages/apps/gui/src/docwen_gui/app.py")
    release = _read("packages/apps/gui/src/docwen_gui/release_smoke.py")
    editor = _read("packages/apps/gui/src/docwen_gui/widgets/settings/toml_editor.py")
    proofread = _read("packages/apps/gui/src/docwen_gui/widgets/settings/proofread_tab.py")
    adapter = _read("packages/bundle/src/docwen_bundle/config_port.py")

    assert port.count("all-or-nothing") >= 6
    assert "crash/power-loss atomicity" in port
    assert "Settings persistence requires transactional ConfigPort.set_many" in settings_vm
    persist = _method_source(
        "packages/apps/gui/src/docwen_gui/view_models/settings_vm.py",
        "SettingsViewModel",
        "_persist_config_values",
    )
    assert "for key, value in values.items()" not in persist
    assert app.count("config_port.set_many(") >= 1
    assert release.count("config_port.set_many(") >= 1
    assert "output_config_persist_failed" in app
    assert "output_config_persist_failed" in release
    assert "atomic_write_text(path, content)" in editor
    assert "if reload_result is False:" in editor
    assert "atomic_write_text(fallback_path, content)" in proofread
    assert "configuration state is untrusted after a failed reconciliation" in adapter
    assert "return self._trusted_config().as_dict()" in adapter
