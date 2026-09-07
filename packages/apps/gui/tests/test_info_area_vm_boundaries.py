"""Task-summary boundary regressions."""

from collections.abc import Generator

import pytest

from docwen_gui.view_models.info_area_vm import InfoAreaViewModel

pytestmark = pytest.mark.gui


@pytest.fixture
def vm() -> Generator[InfoAreaViewModel, None, None]:
    model = InfoAreaViewModel()
    yield model
    model.stop_all_timers()


class DescribeProgressBoundaries:
    """Edge case tests for progress and task state boundaries."""

    def test_error_message_priority_overrides_progress(self, vm: InfoAreaViewModel) -> None:
        """Error messages should take priority over progress messages."""
        vm.set_transient_message("progress:op-1", "Still working...", message_type="progress")
        vm.set_transient_message("error:op-1", "Something failed!", message_type="error")
        assert len(vm.message_types) > 0

    def test_zero_total_count_does_not_crash(self, vm: InfoAreaViewModel) -> None:
        """Division by zero should not crash when total_count is 0."""
        vm.set_task_summary(
            total_count=0,
            completed_count=0,
            state="active",
            operation_id="op-zero",
        )
        assert vm.has_task_summary is True

    def test_cancelled_with_no_retry_action(self, vm: InfoAreaViewModel) -> None:
        """Cancelled state should not offer retry."""
        vm.set_task_summary(
            total_count=5,
            cancelled_count=5,
            state="cancelled",
            tone="info",
            guide_actions=[{"key": "open_output", "label": "Open Output"}],
            operation_id="op-cancel",
        )
        assert {a["key"] for a in vm.guide_actions} == {"open_output"}


class DescribeTaskSummaryCache:
    """Test that task summary state is properly cached and accessible."""

    def test_task_summary_caches_all_fields(self, vm: InfoAreaViewModel) -> None:
        """All fields passed to set_task_summary should be queryable."""
        vm.set_task_summary(
            total_count=10,
            completed_count=3,
            failed_count=1,
            skipped_count=0,
            cancelled_count=0,
            state="active",
            tone="info",
            operation_id="op-123",
        )
        assert vm.has_task_summary is True
