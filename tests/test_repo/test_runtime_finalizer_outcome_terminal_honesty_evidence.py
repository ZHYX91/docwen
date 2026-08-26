from __future__ import annotations

from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]
FINALIZER = ROOT / "packages/runtime/src/docwen_runtime/output/finalizer.py"
TASK_MANAGER = ROOT / "packages/runtime/src/docwen_runtime/engine/task_manager.py"
MAIN_WINDOW = ROOT / "packages/apps/gui/src/docwen_gui/main_window.py"
BATCH_LIST = ROOT / "packages/apps/gui/src/docwen_gui/widgets/batch_list.py"
FINALIZER_TESTS = ROOT / "packages/runtime/tests/test_output_finalizer_*.py"
OUTCOME_TESTS = ROOT / "packages/runtime/tests/test_task_manager_outcome_honesty.py"
NUMBERING_TESTS = ROOT / "packages/runtime/tests/test_request_scoped_numbering.py"
GUI_TESTS = ROOT / "packages/apps/gui/tests/test_main_window_projection_binding_*.py"
BATCH_LIST_TESTS = ROOT / "packages/apps/gui/tests/test_batch_list_widget_*.py"
REPORT_NAME = "runtime-finalizer-outcome-terminal-honesty-2026-07-21.md"


def test_runtime_finalizer_and_terminal_paths_fail_closed() -> None:
    finalizer = FINALIZER.read_text(encoding="utf-8")
    task_manager = TASK_MANAGER.read_text(encoding="utf-8")

    for token in (
        'if overwrite_mode not in {"error", "rename", "overwrite", "skip"}:',
        'if overwrite_mode != "rename" and not io_destination.is_file():',
        "if not cls._io_path(artifact.staging_path).is_file():",
        "cls._ensure_contained(output_dir, item.destination)",
        'summary_code = "FINALIZER_NO_ARTIFACTS"',
        'summary_code = "FINALIZER_PARTIAL" if placed_artifacts else "FINALIZER_FAILED"',
        'error_type="output_finalization_failed"',
        'code="FINALIZER_DONE"',
        "success=error is None",
    ):
        assert token in finalizer

    for token in (
        "if not plugin_result.success or plugin_result.error is not None:",
        'diagnostic_code="PLUGIN_REPORTED_FAILURE"',
        "intentional_empty_success = not artifacts and self._is_intentional_no_output_success(",
        'if request.action_name != "validate":',
        'diagnostic.code == "PROOFREAD-SKIPPED"',
        "if result.success:",
        "terminal_event = make_task_failed(",
        "terminal_diagnostics = self._emit_terminal(state.on_event, terminal_event)",
        'code="TASK_EVENT_LISTENER_ERROR"',
        "runtime_error = ConversionErrorInfo(",
        "**self._metrics_extra(plugin_metrics)",
    ):
        assert token in task_manager


def test_gui_retains_real_failed_artifacts_without_promoting_success() -> None:
    main_window = MAIN_WINDOW.read_text(encoding="utf-8")
    batch_list = BATCH_LIST.read_text(encoding="utf-8")

    for token in (
        'cancelled = bool(error is not None and error.error_type == "cancelled")',
        'retained_output_path = "" if cancelled else self._pick_existing_output_path(result)',
        "primary = [artifact for artifact in result.artifacts if artifact.is_primary]",
        "for artifact in (*primary, *secondary):",
        "filesystem_path(artifact.staging_path).is_file()",
        "except (OSError, ValueError):",
        "output_path=retained_output_path",
        "retained_failure_output_dir",
        "successful_output_dir",
    ):
        assert token in main_window

    assert 'if status in {"completed", "failed"} and entry.output_path:' in batch_list
    assert 'self._primary_action_key = "open_output"' in batch_list
    assert 'self.retry_button.clicked.connect(lambda: self.action_requested.emit("retry_failed"' in batch_list
    assert 'if status == "failed":' in batch_list


def test_regressions_cover_output_truth_terminal_truth_and_failed_artifact_reachability() -> None:
    finalizer_tests = read_source_text(FINALIZER_TESTS)
    outcome_tests = OUTCOME_TESTS.read_text(encoding="utf-8")
    numbering_tests = NUMBERING_TESTS.read_text(encoding="utf-8")
    gui_tests = read_source_text(GUI_TESTS)
    batch_list_tests = read_source_text(BATCH_LIST_TESTS)

    for token in (
        "test_finalize_missing_staging_source_returns_failure_without_phantom_artifact",
        "test_finalize_partial_placement_preserves_real_artifacts_and_reports_typed_error",
        "test_finalize_skip_rejects_existing_directory_as_artifact",
        "test_finalize_overwrite_rejects_existing_directory_as_artifact",
        "test_finalize_rejects_unknown_overwrite_mode_without_mutating_existing_output",
        "test_finalize_empty_artifacts",
    ):
        assert token in finalizer_tests

    for token in (
        "test_reported_plugin_failure_emits_failed_without_spurious_finalizing_progress",
        "test_finalizer_partial_failure_returns_typed_error_and_failed_terminal",
        "test_real_proofread_disabled_result_is_an_intentional_empty_report_success",
        "test_ordinary_success_without_artifacts_is_a_typed_finalizer_failure",
        "test_plugin_success_with_non_cancel_error_is_normalized_to_failed",
        "test_finalizer_exception_preserves_runtime_metrics_and_plugin_extras",
        "test_terminal_listener_rejection_cannot_change_success_or_duplicate_terminal",
    ):
        assert token in outcome_tests

    assert 'diagnostic_code == "PLUGIN_REPORTED_FAILURE"' in numbering_tests
    for token in (
        "test_failed_result_with_retained_auxiliary_exposes_output_without_marking_success",
        "test_batch_failed_result_with_retained_auxiliary_exposes_output_without_changing_counts",
        "test_batch_multiple_retained_failures_keep_each_entry_output_reachable",
        "test_batch_cancelled_or_skipped_result_clears_stale_output",
    ):
        assert token in gui_tests
    assert "test_failed_entry_with_retained_output_can_open_it" in batch_list_tests
