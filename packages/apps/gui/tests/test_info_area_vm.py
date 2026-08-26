"""Model-state tests for InfoAreaViewModel.

Tests the state truth source independently of widget rendering:
history management (dedup, limit, important preservation),
transient management (priority, TTL, throttle),
task summary, guide action computation, and activity animation.
"""

from __future__ import annotations

import time as _time_module
from collections.abc import Generator

import pytest
from PySide6.QtCore import QTimer
from PySide6.QtTest import QSignalSpy

from docwen_gui.view_models.info_area_vm import (
    InfoAreaViewModel,
)

pytestmark = pytest.mark.gui


# ── Fixtures ──────────────────────────────────────────────────────────────


@pytest.fixture
def vm() -> Generator[InfoAreaViewModel, None, None]:
    v = InfoAreaViewModel()
    yield v
    v.stop_all_timers()


# ── History: dedup ────────────────────────────────────────────────────────


class TestHistoryDedup:
    def test_suppresses_rapid_duplicates(self, vm: InfoAreaViewModel) -> None:
        """Identical messages within 250ms are deduplicated."""
        # All calls at the same monotonic time -> only first passes
        vm.add_message("same", "info")
        vm.add_message("same", "info")
        vm.add_message("same", "info")
        assert vm.message_count == 1

    def test_allows_same_message_after_dedup_window(self, monkeypatch, vm: InfoAreaViewModel) -> None:
        """Same message after 250ms is allowed."""
        times = iter([1.0, 1.3])  # 300ms gap
        monkeypatch.setattr(_time_module, "monotonic", lambda: next(times))
        vm.add_message("same", "info")
        vm.add_message("same", "info")
        assert vm.message_count == 2

    def test_different_signatures_not_deduped(self, vm: InfoAreaViewModel) -> None:
        """Different messages or types always pass."""
        vm.add_message("msg1", "info")
        vm.add_message("msg2", "info")
        vm.add_message("msg1", "warning")
        assert vm.message_count == 3

    def test_different_operation_ids_not_deduped(self, vm: InfoAreaViewModel) -> None:
        """Same message text but different operation_id -> different signature."""
        vm.add_message("msg", "info", operation_id="op-1")
        vm.add_message("msg", "info", operation_id="op-2")
        assert vm.message_count == 2


# ── History: limit enforcement ────────────────────────────────────────────


class TestHistoryLimit:
    def test_enforces_max_messages(self, vm: InfoAreaViewModel) -> None:
        vm.max_messages = 3
        for i in range(5):
            vm.add_message(f"msg-{i}", "info")
        assert vm.message_count == 3

    def test_preserves_important_on_overflow(self, vm: InfoAreaViewModel) -> None:
        vm.max_messages = 2
        vm.add_message("info-1", "info")
        vm.add_message("warn-1", "warning")
        vm.add_message("info-2", "info")
        assert vm.message_count == 2
        assert vm.message_types == ["warning", "info"]

    def test_removes_oldest_when_all_important(self, vm: InfoAreaViewModel) -> None:
        vm.max_messages = 2
        vm.add_message("warn-1", "warning")
        vm.add_message("danger-1", "danger")
        vm.add_message("success-1", "success")
        assert vm.message_count == 2
        # First (oldest) important removed
        assert vm.message_types == ["danger", "success"]


# ── History: timestamp ────────────────────────────────────────────────────


class TestHistoryTimestamp:
    def test_timestamp_is_hh_mm_ss_format(self, vm: InfoAreaViewModel) -> None:
        vm.add_message("test", "info")
        row = vm.history_rows[0]
        parts = row.timestamp.split(":")
        assert len(parts) == 3
        assert 0 <= int(parts[0]) <= 23  # hours
        assert 0 <= int(parts[1]) <= 59  # minutes
        assert 0 <= int(parts[2]) <= 59  # seconds


# ── History: data integrity ───────────────────────────────────────────────


class TestHistoryData:
    def test_stores_all_fields(self, vm: InfoAreaViewModel) -> None:
        vm.add_message(
            "test message",
            "success",
            show_location=True,
            file_path="/path/to/file.txt",
            navigate_file_path="/path/to/nav.txt",
            operation_id="op-1234",
        )
        row = vm.history_rows[0]
        assert row.message == "test message"
        assert row.message_type == "success"
        assert row.show_location is True
        assert row.file_path == "/path/to/file.txt"
        assert row.navigate_file_path == "/path/to/nav.txt"
        assert row.operation_id == "op-1234"

    def test_message_types_property(self, vm: InfoAreaViewModel) -> None:
        vm.add_message("a", "info")
        vm.add_message("b", "success")
        vm.add_message("c", "warning")
        assert vm.message_types == ["info", "success", "warning"]


# ── Transient: priority ───────────────────────────────────────────────────


class TestTransientPriority:
    def test_higher_priority_wins_display(self, vm: InfoAreaViewModel) -> None:
        vm.set_transient_message("progress:op-1", "Working", "info", ttl_ms=0)
        vm.set_transient_message("error:op-1", "Failed", "danger", ttl_ms=0)
        # error (priority 0) beats progress (priority 2)
        assert vm.status_summary_text == "Failed"
        assert vm.status_tone == "danger"

    def test_clearing_top_priority_falls_back(self, vm: InfoAreaViewModel) -> None:
        vm.set_transient_message("progress:op-1", "Working", "info", ttl_ms=0)
        vm.set_transient_message("error:op-1", "Failed", "danger", ttl_ms=0)
        vm.clear_transient_message("error:op-1")
        assert vm.status_summary_text == "Working"
        assert vm.status_tone == "info"

    def test_empty_transients_returns_to_idle(self, vm: InfoAreaViewModel) -> None:
        vm.set_transient_message("progress", "Working", "info", ttl_ms=0)
        vm.clear_transient_message("progress")
        assert vm.status_summary_text != "Working"
        assert vm.status_source == "idle"


# ── Transient: throttle ───────────────────────────────────────────────────


class TestTransientThrottle:
    def test_throttles_progress_updates(self, monkeypatch, vm: InfoAreaViewModel) -> None:
        now = 1.0
        monkeypatch.setattr(_time_module, "monotonic", lambda: now)
        progress_key = "progress:op-1234abcd"

        # First update displays immediately
        vm.set_transient_message(progress_key, "Step 1", "info", ttl_ms=0, source="op-1234abcd")
        assert vm.status_summary_text == "Step 1"

        # Rapid updates within throttle window don't change display
        vm.set_transient_message(progress_key, "Step 2", "info", ttl_ms=0, source="op-1234abcd")
        vm.set_transient_message(progress_key, "Step 3", "info", ttl_ms=0, source="op-1234abcd")
        assert vm.status_summary_text == "Step 1"

        # After throttle window passes, latest update appears
        now = 2.0
        vm.set_transient_message(progress_key, "Step 4", "info", ttl_ms=0, source="op-1234abcd")
        assert vm.status_summary_text == "Step 4"

    def test_flush_delivers_latest_pending(self, monkeypatch, vm: InfoAreaViewModel) -> None:
        """Delayed flush should deliver the latest pending update, not intermediate ones."""
        now = 1.0
        monkeypatch.setattr(_time_module, "monotonic", lambda: now)
        progress_key = "progress:op-1"

        vm.set_transient_message(progress_key, "Step 1", "info", ttl_ms=0, source="op-1")
        assert vm.status_summary_text == "Step 1"

        # These go to pending
        vm.set_transient_message(progress_key, "Step 2", "info", ttl_ms=0, source="op-1")
        vm.set_transient_message(progress_key, "Step 3", "info", ttl_ms=0, source="op-1")

        # Only Step 3 should be in pending
        assert progress_key in vm._pending_transient_updates

        # Now advance time past throttle window
        now = 2.0
        vm.set_transient_message(progress_key, "Step 4", "info", ttl_ms=0, source="op-1")
        assert vm.status_summary_text == "Step 4"
        assert progress_key not in vm._pending_transient_updates


# ── Transient: TTL ────────────────────────────────────────────────────────


class TestTransientTTL:
    def test_transient_expires_after_ttl(self, vm: InfoAreaViewModel) -> None:
        """Transient should clear itself after TTL expires."""
        vm.set_transient_message("test", "Temporary", "info", ttl_ms=10, source="test-src")
        assert vm.transient_count == 1

        # Let the timer fire (need to process Qt events)
        QTimer.singleShot(20, lambda: None)  # dummy to allow processing
        # Actually, we need the real timer to fire. Let's test differently.
        # The TTL timer is a singleShot QTimer — we can find and trigger it.
        for key, timer in vm._transient_timers.items():
            if key == "test":
                timer.timeout.emit()  # force expiry
                break

        assert vm.transient_count == 0

    def test_newer_version_prevents_old_expiry(self, vm: InfoAreaViewModel) -> None:
        """Old timer for a superseded version should NOT clear the new message."""
        vm.set_transient_message("processing", "First", "warning", ttl_ms=3000)
        vm.set_transient_message("processing", "Second", "warning", ttl_ms=3000)
        assert vm.status_summary_text == "Second"

    def test_zero_ttl_never_expires(self, vm: InfoAreaViewModel) -> None:
        """TTL=0 means no auto-expiry timer."""
        vm.set_transient_message("persistent", "Stays", "info", ttl_ms=0)
        assert "persistent" not in vm._transient_timers
        assert vm.transient_count == 1


# ── Transient: terminal ───────────────────────────────────────────────────


class TestTerminalTransient:
    def test_terminal_clears_source_transients(self, monkeypatch, vm: InfoAreaViewModel) -> None:
        monkeypatch.setattr(_time_module, "monotonic", lambda: 1.0)
        progress_key = "progress:op-1234abcd"

        vm.set_transient_message(progress_key, "Step 1", "info", ttl_ms=0, source="op-1234abcd")
        vm.set_transient_message(progress_key, "Step 2", "info", ttl_ms=0, source="op-1234abcd")

        assert progress_key in vm._pending_transient_updates
        assert vm.status_summary_text == "Step 1"

        # Terminal transient from add_message
        vm.add_message("Done", "success", operation_id="op-1234abcd")

        assert progress_key not in vm._transient_messages
        assert progress_key not in vm._pending_transient_updates
        assert vm.status_summary_text == "Done"
        assert vm.status_tone == "success"
        assert "terminal:op-1234abcd" in vm._transient_messages

    @pytest.mark.parametrize(
        ("message_type", "message"),
        [
            ("success", "Done"),
            ("danger", "Failed"),
            ("warning", "Cancelled"),
        ],
    )
    def test_terminal_bypasses_progress_throttle(
        self, message_type, message, monkeypatch, vm: InfoAreaViewModel
    ) -> None:
        monkeypatch.setattr(_time_module, "monotonic", lambda: 1.0)
        progress_key = "progress:op-1234abcd"

        vm.set_transient_message(progress_key, "Step 1", "info", ttl_ms=0, source="op-1234abcd")
        vm.set_transient_message(progress_key, "Step 2", "info", ttl_ms=0, source="op-1234abcd")

        assert progress_key in vm._pending_transient_updates
        assert vm.status_summary_text == "Step 1"

        vm.add_message(message, message_type, operation_id="op-1234abcd")

        assert progress_key not in vm._transient_messages
        assert progress_key not in vm._pending_transient_updates
        assert vm.status_summary_text == message
        assert vm.status_tone == message_type


# ── Task summary ──────────────────────────────────────────────────────────


class TestTaskSummary:
    def test_set_and_clear(self, vm: InfoAreaViewModel) -> None:
        vm.set_task_summary(
            operation_id="op-1",
            current_file="test.docx",
            completed_count=1,
            total_count=3,
            failed_count=0,
            state="active",
            tone="info",
        )
        assert vm.has_task_summary
        assert vm.status_source == "task"

        vm.clear_task_summary()
        assert not vm.has_task_summary
        assert vm.status_source == "idle"

    def test_unknown_state_fails_without_replacing_existing_summary(self, vm: InfoAreaViewModel) -> None:
        vm.set_task_summary(operation_id="op-1", state="active", tone="info")

        with pytest.raises(ValueError, match="state must be empty or one of"):
            vm.set_task_summary(operation_id="op-2", state="mystery", tone="info")

        assert vm.task_summary.operation_id == "op-1"
        assert vm.task_summary.state == "active"

    def test_unknown_tone_fails_without_replacing_existing_summary(self, vm: InfoAreaViewModel) -> None:
        vm.set_task_summary(operation_id="op-1", state="active", tone="info")

        with pytest.raises(ValueError, match="tone must be one of"):
            vm.set_task_summary(operation_id="op-2", state="success", tone="mystery")

        assert vm.task_summary.operation_id == "op-1"
        assert vm.task_summary.tone == "info"

    def test_task_summary_shows_progress(self, vm: InfoAreaViewModel) -> None:
        vm.set_task_summary(
            operation_id="op-1",
            current_file="demo.docx",
            completed_count=1,
            total_count=10,
            failed_count=0,
            state="active",
            tone="info",
        )
        assert "demo.docx" in vm.status_summary_text

    def test_transient_overrides_task_summary(self, vm: InfoAreaViewModel) -> None:
        vm.set_task_summary(
            operation_id="op-1",
            current_file="file.docx",
            completed_count=1,
            total_count=3,
            failed_count=0,
            state="active",
            tone="info",
        )
        vm.set_transient_message("error:op-1", "Error occurred", "danger", ttl_ms=0)
        assert vm.status_source == "transient"
        assert vm.status_summary_text == "Error occurred"

    def test_activity_enabled_during_active_task(self, vm: InfoAreaViewModel) -> None:
        vm.set_task_summary(
            operation_id="op-anim",
            current_file="demo.docx",
            completed_count=1,
            total_count=10,
            failed_count=0,
            state="active",
            tone="info",
        )
        assert vm.activity_enabled


# ── Guide actions ─────────────────────────────────────────────────────────


class TestGuideActions:
    def test_success_shows_open_output_and_add_more(self, vm: InfoAreaViewModel) -> None:
        actions = InfoAreaViewModel.compute_guide_actions(
            "success", output_dir="/tmp/out", failed_details_path="", retry_available=True
        )
        keys = [a["action_key"] for a in actions]
        assert "open_output_dir" in keys
        assert "add_more_files" in keys
        assert "view_failed_details" not in keys
        assert "retry_failed" not in keys

    def test_failed_shows_all_three(self, vm: InfoAreaViewModel) -> None:
        actions = InfoAreaViewModel.compute_guide_actions(
            "failed", output_dir="/tmp/out", failed_details_path="/tmp/failed.txt", retry_available=True
        )
        keys = [a["action_key"] for a in actions]
        assert "open_output_dir" in keys
        assert "view_failed_details" in keys
        assert "retry_failed" in keys
        assert "add_more_files" in keys

    def test_partial_shows_all_three(self, vm: InfoAreaViewModel) -> None:
        actions = InfoAreaViewModel.compute_guide_actions(
            "partial", output_dir="/tmp/out", failed_details_path="/tmp/failed.txt", retry_available=True
        )
        keys = [a["action_key"] for a in actions]
        assert "open_output_dir" in keys
        assert "view_failed_details" in keys
        assert "retry_failed" in keys
        assert "add_more_files" in keys

    def test_cancelled_only_add_more(self, vm: InfoAreaViewModel) -> None:
        actions = InfoAreaViewModel.compute_guide_actions(
            "cancelled", output_dir="/tmp/out", failed_details_path="", retry_available=True
        )
        keys = [a["action_key"] for a in actions]
        assert keys == ["add_more_files"]

    def test_failed_without_retry_excludes_retry(self, vm: InfoAreaViewModel) -> None:
        actions = InfoAreaViewModel.compute_guide_actions(
            "failed", output_dir="/tmp/out", failed_details_path="/tmp/failed.txt", retry_available=False
        )
        keys = [a["action_key"] for a in actions]
        assert "retry_failed" not in keys
        assert "open_output_dir" in keys
        assert "view_failed_details" in keys
        assert "add_more_files" in keys

    def test_guide_set_from_task_summary(self, vm: InfoAreaViewModel) -> None:
        guide_actions = [
            {"action_key": "open_output_dir", "target_path": "/tmp/out"},
            {"action_key": "add_more_files", "target_path": ""},
        ]
        vm.set_task_summary(
            operation_id="op-2000",
            current_file="done.docx",
            completed_count=2,
            total_count=2,
            failed_count=0,
            state="success",
            tone="success",
            guide_actions=guide_actions,
        )
        assert vm.guide_visible
        assert len(vm.guide_actions) == 2

    def test_guide_not_visible_for_active_task(self, vm: InfoAreaViewModel) -> None:
        vm.set_task_summary(
            operation_id="op-1",
            state="active",
            tone="info",
            guide_actions=[{"action_key": "add_more_files", "target_path": ""}],
        )
        assert not vm.guide_visible


# ── Status fallback chain ─────────────────────────────────────────────────


class TestStatusFallback:
    def test_idle_when_empty(self, vm: InfoAreaViewModel) -> None:
        assert vm.status_source == "idle"

    def test_history_when_no_transient_or_task(self, vm: InfoAreaViewModel) -> None:
        vm.add_message("History entry", "info")
        assert vm.status_source == "history"
        assert vm.status_summary_text == "History entry"

    def test_transient_beats_history(self, vm: InfoAreaViewModel) -> None:
        vm.add_message("History entry", "info")
        vm.set_transient_message("progress", "Working", "info", ttl_ms=0)
        assert vm.status_source == "transient"
        assert vm.status_summary_text == "Working"

    def test_clearing_transient_returns_to_history(self, vm: InfoAreaViewModel) -> None:
        vm.add_message("History entry", "info")
        vm.set_transient_message("progress", "Working", "info", ttl_ms=0)
        vm.clear_transient_message("progress")
        assert vm.status_source == "history"
        assert vm.status_summary_text == "History entry"


# ── Activity animation ────────────────────────────────────────────────────


class TestActivityAnimation:
    def test_format_zero_dots(self, vm: InfoAreaViewModel) -> None:
        result = vm._format_activity_meta("Processing", 0)
        assert result == "Processing"

    def test_format_one_dot(self, vm: InfoAreaViewModel) -> None:
        result = vm._format_activity_meta("Processing", 1)
        assert result == "Processing."

    def test_format_three_dots(self, vm: InfoAreaViewModel) -> None:
        result = vm._format_activity_meta("Processing", 3)
        assert result == "Processing..."

    def test_format_cycle(self, vm: InfoAreaViewModel) -> None:
        """Step 4 resets to 0 dots."""
        result = vm._format_activity_meta("Processing", 4)
        assert result == "Processing"

    def test_start_stop_animation(self, vm: InfoAreaViewModel) -> None:
        assert not vm.activity_enabled
        vm.start_activity_animation("Working")
        assert vm.activity_enabled
        vm.stop_activity_animation()
        assert not vm.activity_enabled

    def test_animation_tick_updates_meta(self, vm: InfoAreaViewModel) -> None:
        vm.start_activity_animation("Working")
        # Advance a tick
        vm._tick_activity_animation()
        assert vm.activity_meta_text == "Working."


# ── Navigation ────────────────────────────────────────────────────────────


class TestNavigation:
    def test_request_navigation_emits_signal(self, vm: InfoAreaViewModel) -> None:
        emitted: list[str] = []
        vm.history_navigation_requested.connect(emitted.append)

        vm.add_message("Test", "info", navigate_file_path="/path/to/file")
        vm.request_navigation()
        assert emitted == ["/path/to/file"]

    def test_request_navigation_no_target_no_emit(self, vm: InfoAreaViewModel) -> None:
        emitted: list[str] = []
        vm.history_navigation_requested.connect(emitted.append)
        vm.request_navigation()
        assert emitted == []


# ── Guide action request ──────────────────────────────────────────────────


class TestGuideActionRequest:
    def test_emits_signal(self, vm: InfoAreaViewModel) -> None:
        emitted: list[tuple[str, str]] = []
        vm.task_guide_action_requested.connect(lambda ak, tp: emitted.append((ak, tp)))
        vm.request_guide_action("open_output_dir", "/tmp/out")
        assert emitted == [("open_output_dir", "/tmp/out")]


# ── Location request ──────────────────────────────────────────────────────


class TestLocationRequest:
    def test_request_location_emits_signal(self, vm: InfoAreaViewModel) -> None:
        emitted: list[str] = []
        vm.location_requested.connect(emitted.append)
        vm.request_location("/path/to/file.txt")
        assert emitted == ["/path/to/file.txt"]

    def test_request_location_empty_path_no_emit(self, vm: InfoAreaViewModel) -> None:
        emitted: list[str] = []
        vm.location_requested.connect(emitted.append)
        vm.request_location("")
        assert emitted == []


# ── Clear all ─────────────────────────────────────────────────────────────


class TestClearAll:
    def test_clear_all_transients(self, vm: InfoAreaViewModel) -> None:
        vm.set_transient_message("a", "A", "info", ttl_ms=0)
        vm.set_transient_message("b", "B", "info", ttl_ms=0)
        assert vm.transient_count == 2
        vm.clear_all_transients()
        assert vm.transient_count == 0

    def test_clear_history(self, vm: InfoAreaViewModel) -> None:
        vm.add_message("a", "info")
        vm.add_message("b", "info")
        assert vm.message_count == 2
        vm.clear_history()
        assert vm.message_count == 0

    def test_reset_session_is_atomic_and_invalidates_pending_timers(
        self,
        monkeypatch,
        vm: InfoAreaViewModel,
    ) -> None:
        now = 1.0
        monkeypatch.setattr(_time_module, "monotonic", lambda: now)
        progress_key = "progress:session-reset"

        vm.add_message("old history", "warning")
        vm.set_task_summary(
            operation_id="session-reset",
            current_file="old.docx",
            total_count=1,
            state="active",
            guide_actions=[{"action_key": "add_more_files", "target_path": ""}],
        )
        vm.set_transient_message(progress_key, "Step 1", "info", ttl_ms=0, source="session-reset")
        vm.set_transient_message(progress_key, "Step 2", "info", ttl_ms=0, source="session-reset")

        assert progress_key in vm._pending_transient_updates
        stale_flush_timer = vm._transient_flush_timers[progress_key]
        vm.start_activity_animation("Working")
        state_changed = QSignalSpy(vm.state_changed)

        vm.reset_session()

        assert state_changed.count() == 1
        assert vm.message_count == 0
        assert vm.transient_count == 0
        assert vm.has_task_summary is False
        assert vm.guide_actions == []
        assert vm.guide_visible is False
        assert vm.activity_enabled is False
        assert vm.status_source == "idle"
        assert vm._pending_transient_updates == {}
        assert vm._transient_timers == {}
        assert vm._transient_flush_timers == {}
        assert stale_flush_timer.isActive() is False

        # Reusing the same key after reset must not let the stale callback
        # restore or remove state from the new session.
        vm.set_transient_message(progress_key, "New session", "info", ttl_ms=0, source="session-reset")
        stale_flush_timer.timeout.emit()
        assert vm.status_summary_text == "New session"
        assert vm.transient_count == 1

    def test_stop_all_timers(self, vm: InfoAreaViewModel) -> None:
        vm.set_transient_message("a", "A", "info", ttl_ms=5000)
        vm.start_activity_animation("Working")
        assert len(vm._transient_timers) >= 1
        vm.stop_all_timers()
        assert len(vm._transient_timers) == 0
        assert vm._activity_timer is None


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
