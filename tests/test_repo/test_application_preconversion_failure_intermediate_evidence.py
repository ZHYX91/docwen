from __future__ import annotations

from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]
TASK_MANAGER = ROOT / "packages/runtime/src/docwen_runtime/engine/task_manager.py"
RUNTIME_TESTS = ROOT / "packages/runtime/tests/test_fake_closed_loop_*.py"
REPORT_NAME = "application-preconversion-failure-intermediate-preservation-2026-07-21.md"


def test_failure_intermediate_runtime_contract_is_single_shot_and_cancellation_safe() -> None:
    source = TASK_MANAGER.read_text(encoding="utf-8")

    for token in (
        "finalization_attempted: bool = False",
        "plugin_reported_cancellation",
        "token.is_cancelled or plugin_reported_cancellation",
        "_finalize_failure_intermediates(",
        "PRECONVERSION_INTERMEDIATE_FINALIZE_ERROR",
        "TASK_EVENT_LISTENER_ERROR",
        "extra=self._metrics_extra(state.plugin_result.metrics)",
        "output_bytes=0",
    ):
        assert token in source

    cancellation_check = source.index("if token.is_cancelled or plugin_reported_cancellation:")
    finalizing_progress = source.index('"Finalizing output"', cancellation_check)
    failure_branch = source.index(
        "if not plugin_result.success or plugin_result.error is not None:",
        cancellation_check,
    )
    failure_attempt = source.index("state.finalization_attempted = True", failure_branch)
    failure_finalize = source.index("preserved = self._finalize_failure_intermediates(", failure_branch)
    assert cancellation_check < failure_branch < finalizing_progress
    assert finalizing_progress < failure_attempt < failure_finalize


def test_failure_intermediate_regressions_lock_failure_cancel_metrics_and_listener_edges() -> None:
    tests = read_source_text(RUNTIME_TESTS)

    for token in (
        "test_preconversion_intermediate_is_finalized_when_plugin_fails",
        "test_failure_intermediate_placement_error_does_not_mask_plugin_failure",
        "test_structured_plugin_cancellation_does_not_finalize_preconversion_intermediate",
        "test_runtime_token_cancellation_discards_intermediate_despite_plugin_failure",
        "test_failure_intermediate_finalizes_once_and_listener_cannot_mask_plugin_error",
        "assert result.metrics.output_bytes == 0",
        "assert finalize_calls == [request.request_id]",
        '"TASK_EVENT_LISTENER_ERROR"',
        'assert not list(tmp_path.glob("docwen_pre_*"))',
    ):
        assert token in tests
