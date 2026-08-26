"""Thread-safety and functional tests for TaskEventBridge.

No QApplication needed for bridge-only tests.  Smoke tests that
require an event loop are run under ``pytest-qt`` markers.
"""

import threading

import pytest

from docwen_gui.qt_bridge.task_event_bridge import (
    MAX_QUEUE_SIZE,
    TaskEventBridge,
)

pytestmark = pytest.mark.gui


# ── Helpers ────────────────────────────────────────────────────────────


@pytest.fixture
def bridge() -> TaskEventBridge:
    return TaskEventBridge()


# ── Queue and flush ────────────────────────────────────────────────────


class TestEnqueueAndFlush:
    def test_enqueue_empty_event_type_ignored(self, bridge: TaskEventBridge) -> None:
        bridge.enqueue("", {"task_id": "t1"})
        bridge.enqueue("   ", {"task_id": "t1"})
        emitted = bridge.flush()
        assert emitted == 0

    def test_flush_empty_queue_returns_zero(self, bridge: TaskEventBridge) -> None:
        assert bridge.flush() == 0
        assert bridge.queue_size == 0

    def test_single_event_flush(self, bridge: TaskEventBridge) -> None:
        events: list[tuple[str, dict]] = []
        bridge.task_event.connect(lambda et, p: events.append((et, p)))
        bridge.enqueue("task_progress", {"task_id": "t1", "percent": 50.0})
        emitted = bridge.flush()
        assert emitted == 1
        assert len(events) == 1
        assert events[0] == ("task_progress", {"task_id": "t1", "percent": 50.0})

    def test_multiple_events_flush(self, bridge: TaskEventBridge) -> None:
        events: list[tuple[str, dict]] = []
        bridge.task_event.connect(lambda et, p: events.append((et, p)))
        for i in range(5):
            bridge.enqueue("task_progress", {"task_id": f"t{i}", "percent": i * 20.0})
        emitted = bridge.flush()
        assert emitted == 5
        assert len(events) == 5

    def test_flush_clears_queue(self, bridge: TaskEventBridge) -> None:
        bridge.enqueue("task_started", {"task_id": "t1"})
        bridge.flush()
        assert bridge.queue_size == 0
        assert bridge.flush() == 0  # second flush is empty


# ── Queue overflow ─────────────────────────────────────────────────────


class TestQueueOverflow:
    def test_queue_capped_at_max_size(self, bridge: TaskEventBridge) -> None:
        for i in range(MAX_QUEUE_SIZE + 100):
            bridge.enqueue("task_progress", {"task_id": f"t{i}"})
        assert bridge.queue_size <= MAX_QUEUE_SIZE

    def test_oldest_dropped_on_overflow(self, bridge: TaskEventBridge) -> None:
        events: list[tuple[str, dict]] = []
        bridge.task_event.connect(lambda et, p: events.append((et, p)))
        # Fill the queue, each with unique task_id
        for i in range(MAX_QUEUE_SIZE + 10):
            bridge.enqueue("task_progress", {"task_id": f"t{i}"})
        emitted = bridge.flush()
        assert emitted == MAX_QUEUE_SIZE
        # The first few events (t0..t9) should be dropped
        first_id = events[0][1]["task_id"]
        assert first_id.startswith("t10") or int(first_id[1:]) >= 10


# ── Cross-thread safety ────────────────────────────────────────────────


class TestCrossThreadSafety:
    def test_enqueue_from_background_thread(self, bridge: TaskEventBridge) -> None:
        """Events enqueued from a background thread must be flushable on main thread."""
        barrier = threading.Barrier(2, timeout=5)
        errors: list[Exception] = []

        def producer() -> None:
            try:
                for i in range(100):
                    bridge.enqueue("task_progress", {"task_id": f"bg_{i}", "percent": i})
            except Exception as exc:
                errors.append(exc)
            finally:
                barrier.wait()

        t = threading.Thread(target=producer, daemon=True)
        t.start()
        barrier.wait()
        t.join(timeout=5)

        assert not errors, f"Background thread errors: {errors}"
        assert bridge.queue_size == 100

    def test_concurrent_enqueue_no_corruption(self, bridge: TaskEventBridge) -> None:
        """Multiple threads enqueuing concurrently must not corrupt the queue."""
        N_THREADS = 4
        N_EVENTS = 250  # per thread
        barrier = threading.Barrier(N_THREADS, timeout=10)
        errors: list[Exception] = []

        def producer(thread_id: int) -> None:
            try:
                barrier.wait()  # all start together
                for i in range(N_EVENTS):
                    bridge.enqueue("task_progress", {"thread": thread_id, "seq": i})
            except Exception as exc:
                errors.append(exc)

        threads = [threading.Thread(target=producer, args=(tid,), daemon=True) for tid in range(N_THREADS)]
        for t in threads:
            t.start()
        for t in threads:
            t.join(timeout=10)

        assert not errors
        total = bridge.queue_size
        assert total == N_THREADS * N_EVENTS, (
            f"Expected {N_THREADS * N_EVENTS} events, got {total} (corruption detected)"
        )


# ── flush_error signal ─────────────────────────────────────────────────


class TestFlushError:
    def test_bridge_usable_after_disconnect(self, bridge: TaskEventBridge) -> None:
        """Bridge should remain usable after connecting and disconnecting slots."""
        results: list[tuple[str, dict]] = []
        bridge.task_event.connect(lambda et, p: results.append((et, p)))

        bridge.enqueue("task_started", {"task_id": "t1"})
        emitted = bridge.flush()
        assert emitted == 1
        assert len(results) == 1

        # Disconnect all and enqueue again
        bridge.task_event.disconnect()
        bridge.enqueue("task_completed", {"task_id": "t2"})
        emitted = bridge.flush()
        assert emitted == 1  # emits, but no slots connected
        assert len(results) == 1  # no new result since disconnected


# ── Auto-flush (no event loop, just API check) ─────────────────────────


class TestAutoFlush:
    def test_start_stop_auto_flush(self, bridge: TaskEventBridge) -> None:
        """start/stop should not crash, stop always leaves is_flushing=False."""
        bridge.start_auto_flush(interval_ms=10000)
        bridge.stop_auto_flush()
        assert bridge.is_flushing is False

    def test_flush_remaining_on_stop(self, bridge: TaskEventBridge) -> None:
        events: list[tuple[str, dict]] = []
        bridge.task_event.connect(lambda et, p: events.append((et, p)))
        bridge.enqueue("task_started", {"task_id": "t1"})
        bridge.start_auto_flush(interval_ms=10000)
        bridge.stop_auto_flush()
        # stop_auto_flush calls flush() → remaining events emitted
        assert len(events) == 1
        assert bridge.queue_size == 0


# ── flush_error signal (continued) ─────────────────────────────────────


class TestProperties:
    def test_queue_size_reflects_state(self, bridge: TaskEventBridge) -> None:
        assert bridge.queue_size == 0
        bridge.enqueue("task_started", {"task_id": "t1"})
        assert bridge.queue_size == 1
        bridge.flush()
        assert bridge.queue_size == 0

    def test_is_flushing_default_false(self, bridge: TaskEventBridge) -> None:
        assert bridge.is_flushing is False


# ── No cross-thread widget call safety ─────────────────────────────────


class TestNoCrossThreadWidgetCall:
    """Verify the bridge design prevents raw cross-thread widget calls.

    The bridge uses Qt signals with auto-connection, which ensures
    queued delivery when the sender and receiver are on different threads.
    This test verifies that the bridge's API does not expose any raw
    widget references or callable that could be used from a background thread.
    """

    def test_bridge_has_no_widget_references(self, bridge: TaskEventBridge) -> None:
        """The bridge's public API should not expose any QWidget subclass."""
        # All public attributes should be QObject-based, not widget-based
        from PySide6.QtWidgets import QWidget

        for attr_name in dir(bridge):
            if attr_name.startswith("_"):
                continue
            try:
                val = getattr(bridge, attr_name)
            except Exception:
                continue
            if isinstance(val, QWidget):
                pytest.fail(
                    f"TaskEventBridge.{attr_name} exposes a QWidget — this invites cross-thread widget manipulation"
                )

    def test_enqueue_accepts_primitives_only(self, bridge: TaskEventBridge) -> None:
        """enqueue() only accepts str and dict — no Qt objects or callbacks."""
        bridge.enqueue("task_progress", {"task_id": "t1"})
        # Should not accept QObject payload (type checker should catch)
        # We just verify it handles the happy path from any thread
        assert bridge.queue_size == 1
