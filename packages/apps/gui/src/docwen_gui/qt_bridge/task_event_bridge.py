"""TaskEventBridge — thread-safe bridge for runtime task events.

Background threads (e.g. worker pools, plugin execution) emit task
events.  Those events must reach the GUI (widgets) on the *main* thread
without raw cross-thread Qt widget calls.

Design:
    - Producer (any thread) calls ``enqueue(event)``.
    - A lock-protected internal queue buffers events.
    - ``flush()`` (on the main thread, periodically) emits events via
      Qt signals with ``Qt.AutoConnection`` → same-thread → safe.
    - A ``QTimer`` can drive periodic auto-flush.
"""

from __future__ import annotations

import logging
from typing import Any
from typing import cast as _cast

from PySide6.QtCore import (
    QMutex,
    QMutexLocker,
    QObject,
    QTimer,
    Signal,
    Slot,
)

logger = logging.getLogger(__name__)

# Default flush interval (ms)
DEFAULT_FLUSH_INTERVAL_MS = 50

# Maximum queue size (drop oldest on overflow)
MAX_QUEUE_SIZE = 1000


class TaskEventBridge(QObject):
    """Thread-safe bridge from runtime task events to Qt main thread.

    Usage::

        bridge = TaskEventBridge(parent=window)
        bridge.task_event.connect(view_model.on_task_event)

        # In background thread:
        bridge.enqueue("task_progress", {"task_id": "t1", "percent": 50.0})

        # Auto-flush (every 50 ms on main thread):
        bridge.start_auto_flush()
    """

    # ── Signals ────────────────────────────────────────────────────────

    task_event = Signal(str, dict)
    """Emitted on the main thread: (event_type, payload).

    Connect this to ``MainWindowViewModel.on_task_event``.
    """

    flush_error = Signal(str)
    """Emitted when a flush error occurs (e.g. corrupted event)."""

    # ── Construction ────────────────────────────────────────────────────

    def __init__(self, parent: QObject | None = None) -> None:
        super().__init__(parent)
        self._queue: list[tuple[str, dict[str, Any]]] = []
        self._mutex = QMutex()
        self._auto_flush_timer: QTimer = _cast(QTimer, None)
        self._flush_interval_ms = DEFAULT_FLUSH_INTERVAL_MS

    # ── Public API ─────────────────────────────────────────────────────

    def enqueue(self, event_type: str, payload: dict[str, Any]) -> None:
        """Enqueue a task event from any thread.

        Thread-safe.  Can be called from background worker threads.

        Args:
            event_type: One of the ``task_events`` constants
                (``task_started``, ``task_progress``, etc.).
            payload: Event-specific data dict.
        """
        if not event_type.strip():
            logger.debug("TaskEventBridge: ignoring enqueue with empty event_type")
            return

        with QMutexLocker(self._mutex):
            if len(self._queue) >= MAX_QUEUE_SIZE:
                self._queue.pop(0)  # drop oldest
                logger.debug(
                    "TaskEventBridge: queue full (%d), dropped oldest event",
                    MAX_QUEUE_SIZE,
                )
            self._queue.append((event_type, payload))

    def flush(self) -> int:
        """Flush queued events (call on main thread only).

        Returns:
            Number of events emitted.
        """
        # Snapshot the queue under lock, then emit outside the lock
        # to avoid deadlock if a connected slot calls back into enqueue/flush.
        with QMutexLocker(self._mutex):
            if not self._queue:
                return 0
            snapshot = self._queue[:]
            self._queue.clear()

        emitted = 0
        for event_type, payload in snapshot:
            try:
                self.task_event.emit(event_type, payload)
                emitted += 1
            except Exception:
                logger.exception("TaskEventBridge: error emitting event type=%s", event_type)
                self.flush_error.emit(f"Failed to emit {event_type}")

        return emitted

    @Slot()
    def _on_flush_timer(self) -> None:
        """Timer slot — called periodically on the main thread."""
        try:
            self.flush()
        except Exception:
            logger.exception("TaskEventBridge: flush timer error")

    def start_auto_flush(self, interval_ms: int | None = None) -> None:
        """Start periodic auto-flush via QTimer.

        Must be called from the main thread.

        Args:
            interval_ms: Flush interval in ms (default: 50).
        """
        if interval_ms is not None:
            self._flush_interval_ms = max(10, interval_ms)

        if self._auto_flush_timer is not None:
            self._auto_flush_timer.stop()

        self._auto_flush_timer = QTimer(self)
        self._auto_flush_timer.setInterval(self._flush_interval_ms)
        self._auto_flush_timer.timeout.connect(self._on_flush_timer)
        self._auto_flush_timer.start()

    def stop_auto_flush(self) -> None:
        """Stop the auto-flush timer and flush remaining events."""
        if self._auto_flush_timer is not None:
            self._auto_flush_timer.stop()
            self._auto_flush_timer = None  # pyright: ignore[reportAttributeAccessIssue]
        self.flush()  # final flush of remaining events

    # ── Properties ─────────────────────────────────────────────────────

    @property
    def queue_size(self) -> int:
        """Current number of queued events (thread-safe)."""
        with QMutexLocker(self._mutex):
            return len(self._queue)

    @property
    def is_flushing(self) -> bool:
        """Whether auto-flush is active."""
        return self._auto_flush_timer is not None and self._auto_flush_timer.isActive()


__all__ = [
    "DEFAULT_FLUSH_INTERVAL_MS",
    "MAX_QUEUE_SIZE",
    "TaskEventBridge",
]
