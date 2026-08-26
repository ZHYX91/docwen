"""EventAdapter — bridges non-Qt application events to Qt signals.

This adapter listens to ``ApplicationController`` events (e.g. batch
started, files added, config changed) and translates them to Qt signals
that widgets can bind to.  It runs entirely on the main thread.

Thread-safety: all signal emissions happen on the main thread.
The adapter itself is a ``QObject`` and signals use auto-connection
(defaults to ``Qt.AutoConnection`` which chooses ``DirectConnection``
for same-thread and ``QueuedConnection`` for cross-thread).
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from PySide6.QtCore import QObject, Signal

if TYPE_CHECKING:
    from docwen_application.controller import ApplicationController


class EventAdapter(QObject):
    """Adapts application-level events to Qt signals.

    Usage::

        adapter = EventAdapter(controller, parent=window)
        adapter.files_added.connect(view_model.add_files_handler)
        adapter.batch_completed.connect(view_model.on_batch_completed)
    """

    # ── Signals ────────────────────────────────────────────────────────

    files_added = Signal(list)
    """Emitted when files are added (payload: list[FileRef] or list[str])."""

    files_dropped = Signal(list)
    """Emitted when files are dropped via drag-and-drop (GUI-only)."""

    batch_started = Signal(dict)
    """Emitted when a batch operation starts."""

    batch_completed = Signal(dict)
    """Emitted when a batch operation completes."""

    batch_cancelled = Signal(dict)
    """Emitted when a batch operation is cancelled."""

    task_enqueued = Signal(dict)
    """Emitted when a task is enqueued."""

    task_finished = Signal(dict)
    """Emitted when a task finishes."""

    config_changed = Signal(dict)
    """Emitted when configuration changes."""

    # ── GUI-only signals ───────────────────────────────────────────────

    ipc_file_received = Signal(str)
    """Emitted when a file is received via IPC from a second instance."""

    window_activated = Signal()
    """Emitted when the existing window should be brought to front."""

    # ── Construction ────────────────────────────────────────────────────

    def __init__(
        self,
        controller: ApplicationController | None = None,
        parent: QObject | None = None,
    ) -> None:
        super().__init__(parent)
        self._controller = controller

    @property
    def controller(self) -> ApplicationController | None:
        """The injected ApplicationController."""
        return self._controller

    # ── Fire hooks (called by external code on the main thread) ─────────

    def on_files_added(self, files: list[Any]) -> None:
        """Notify that files were added."""
        self.files_added.emit(files)

    def on_files_dropped(self, paths: list[str]) -> None:
        """Notify that files were dropped (GUI drag-and-drop)."""
        self.files_dropped.emit(paths)

    def on_batch_started(self, summary: dict[str, Any]) -> None:
        """Notify that a batch operation has started."""
        self.batch_started.emit(summary)

    def on_batch_completed(self, summary: dict[str, Any]) -> None:
        """Notify that a batch operation has completed."""
        self.batch_completed.emit(summary)

    def on_task_finished(self, result: dict[str, Any]) -> None:
        """Notify that a single task has finished."""
        self.task_finished.emit(result)

    def on_ipc_file_received(self, file_path: str) -> None:
        """Notify that a file was received via IPC."""
        self.ipc_file_received.emit(file_path)

    def on_window_activated(self) -> None:
        """Notify that the existing window should be brought to front."""
        self.window_activated.emit()


__all__ = ["EventAdapter"]
