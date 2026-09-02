"""Evidence guards for VIS-2026-07-22-169 durable config recovery."""

from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]
REPORT_NAME = "config-durable-transaction-recovery-2026-07-22.md"


def _read(relative_path: str) -> str:
    return (PROJECT_ROOT / relative_path).read_text(encoding="utf-8")


def test_runtime_owns_one_locked_checksummed_recovery_protocol() -> None:
    loader = _read("packages/runtime/src/docwen_runtime/config/loader.py")
    transaction = _read("packages/runtime/src/docwen_runtime/config/transaction.py")
    toml_io = _read("packages/runtime/src/docwen_runtime/toml_io.py")

    for token in (
        "config_transaction.process_config_lock(self._user_dir)",
        "config_transaction.recover_transaction_journal(",
        "config_transaction.write_transaction_journal(",
        "config_transaction.mark_transaction_committed(",
        "config_transaction.remove_transaction_journal(",
        "durable_unlink(user_path)",
    ):
        assert token in loader
    for token in (
        'CONFIG_JOURNAL_NAME = ".docwen-config.transaction.json"',
        'CONFIG_LOCK_NAME = ".docwen-config.lock"',
        "class UserFilePreimage:",
        "hashlib.sha256(_canonical_json(payload)).hexdigest()",
        "msvcrt.locking(descriptor, msvcrt.LK_LOCK, 1)",
        "fcntl.flock(descriptor, fcntl.LOCK_EX)",
        "if record.state == _COMMITTED:",
        "for preimage in reversed(record.preimages):",
        "logical not in allowed_paths",
    ):
        assert token in transaction
    for token in (
        "def _sync_windows_directory(",
        "file_flag_backup_semantics = 0x02000000",
        "flush_file_buffers(handle)",
        "os.fsync(descriptor)",
        "_sync_directory(file_path.parent)",
        "def durable_unlink(",
    ):
        assert token in toml_io


def test_vis169_process_and_fault_rows_remain_directly_discoverable() -> None:
    tests = _read("packages/runtime/tests/test_config_durable_transaction_recovery.py")
    for name in (
        "test_atomic_write_flushes_parent_directory_and_preserves_mode",
        "test_handled_multifile_failure_restores_regular_file_metadata",
        "test_process_death_after_first_mutation_recovers_prepared_generation",
        "test_process_death_before_commit_marker_recovers_old_generation",
        "test_process_death_after_commit_marker_keeps_new_generation",
        "test_corrupt_journal_fails_loader_closed_without_deleting_evidence",
        "test_valid_checksum_but_invalid_journal_schema_fails_closed",
        "test_real_local_directory_flush_primitive_succeeds",
        "test_journal_parent_barrier_failure_precedes_every_user_mutation",
        "test_user_file_parent_barrier_failure_compensates_prepared_transaction",
        "test_delete_parent_barrier_failure_restores_deleted_override",
        "test_commit_marker_failure_rolls_back_old_generation",
        "test_committed_cleanup_failure_is_safe_success_and_retried_on_reload",
        "test_prepared_recovery_failure_leaves_journal_for_idempotent_retry",
        "test_two_real_processes_preserve_disjoint_same_file_updates",
        "test_real_reload_blocks_until_multifile_writer_releases_process_lock",
    ):
        assert f"def {name}(" in tests
