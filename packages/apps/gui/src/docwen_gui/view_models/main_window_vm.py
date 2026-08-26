"""MainWindowViewModel — observable state and user-action delegation.

This is the single source of truth for the main window's observable
state.  Widgets bind to its signals and properties; user actions flow
through method calls that delegate to ``ApplicationController``.

Thread-safety: all properties that can be read from a non-GUI thread
are protected by ``QMutex`` or use ``Signal`` with ``Qt.QueuedConnection``.
"""

from __future__ import annotations

import logging
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, Any

from PySide6.QtCore import QMutex, QMutexLocker, QObject, Signal

from docwen_core.models import FILE_INSPECTION_METADATA_KEY
from docwen_gui.file_admission_i18n import render_file_inspection_message
from docwen_gui.i18n import t as _t

logger = logging.getLogger(__name__)

if TYPE_CHECKING:
    from docwen_application.controller import ApplicationController
    from docwen_core.models import FileInspection
    from docwen_core.models.file_ref import FileRef

    from .interaction import MainWindowUiProjection


# ── Status priority levels (matches GUI行为与交互规范.md) ────────────
STATUS_IDLE = "idle"
STATUS_PROCESSING = "processing"
STATUS_COMPLETED = "completed"
STATUS_FAILED = "failed"
STATUS_CANCELLED = "cancelled"

# Default mode (from config, falling back to "single")
DEFAULT_MODE = "single"


@dataclass(frozen=True)
class FileAddOutcome:
    """Files admitted into GUI state and files rejected by core inspection."""

    added: tuple[FileRef, ...]
    rejected: tuple[tuple[str, str], ...]


def _read_default_mode(controller: Any | None) -> str:
    if controller is None:
        return DEFAULT_MODE
    cfg = getattr(controller, "config_port", None)
    if cfg is None:
        return DEFAULT_MODE
    try:
        mode = cfg.get("gui.window.default_mode", DEFAULT_MODE)
        if mode in ("single", "batch"):
            return mode
    except Exception:
        return DEFAULT_MODE
    return DEFAULT_MODE


class MainWindowViewModel(QObject):
    """Observable state for the main window.

    Signals:
        title_changed: emitted when the window title should change.
        status_message_changed: emitted when the status bar message changes.
        task_summary_changed: emitted when a task summary is available.
        files_changed: emitted when the file list changes.
        mode_changed: emitted when single/batch mode changes.
    """

    # ── Signals ────────────────────────────────────────────────────────

    title_changed = Signal(str)
    """Emitted when the window title changes (e.g. task count)."""

    status_message_changed = Signal(str)
    """Emitted when the status bar / info area message changes."""

    task_summary_changed = Signal(dict)
    """Emitted when a task summary is available (payload: dict)."""

    files_changed = Signal(list)
    """Emitted when the file list changes (payload: list[FileRef])."""

    files_cleared = Signal()
    """Emitted when all files are removed."""

    mode_changed = Signal(str)
    """Emitted when single/batch mode changes."""

    shutdown_requested = Signal()
    """Emitted when the user requests application shutdown."""

    window_activation_requested = Signal()
    """Emitted when an external process asks this window to come to the front."""

    ipc_file_received = Signal(str)
    """Emitted when a file path is received via IPC from another instance.

    The payload is the absolute file path.
    """

    ui_projection_changed = Signal(object)
    """Emitted when the main-window UI projection changes.

    Payload is a :class:`~docwen_gui.view_models.interaction.MainWindowUiProjection`.
    The MainWindow binds this to widget visibility and the right-panel stack;
    it must not re-derive business state itself.
    """

    # ── Construction ────────────────────────────────────────────────────

    def __init__(
        self,
        controller: ApplicationController | None = None,
        parent: QObject | None = None,
        *,
        file_inspector: Callable[[str], FileInspection] | None = None,
    ) -> None:
        super().__init__(parent)
        self._controller = controller
        if file_inspector is None:
            from docwen_core.detection import inspect_file

            file_inspector = inspect_file
        self._file_inspector = file_inspector
        self._mutex = QMutex()
        self._files: list[FileRef] = []
        self._mode: str = _read_default_mode(controller)
        self._status_message: str = _t("common.ready", "Ready")
        self._title: str = "DocWen Offline"
        self._current_task_id: str | None = None
        self._active_execution_id: str | None = None
        self._accepted_runtime_task_ids: frozenset[str] = frozenset()
        self._ended_runtime_task_ids: set[str] = set()
        # Selected-file source state — the single source of truth for which
        # file drives the right-panel projection.  Held as the raw FileRef so
        # the projection can read category/format/path without re-deriving.
        self._selected_file: FileRef | None = None

    # ── Observable properties ───────────────────────────────────────────

    @property
    def controller(self) -> ApplicationController | None:
        """The injected ApplicationController (may be None in limited mode)."""
        return self._controller

    @property
    def mode(self) -> str:
        """Current mode: ``"single"`` or ``"batch"``."""
        with QMutexLocker(self._mutex):
            return self._mode

    @mode.setter
    def mode(self, value: str) -> None:
        if value not in ("single", "batch"):
            raise ValueError(f"Invalid mode: {value!r}")
        changed = False
        with QMutexLocker(self._mutex):
            if self._mode != value:
                self._mode = value
                changed = True
        if changed:
            self.mode_changed.emit(value)
        # Recompute projection outside the mutex to avoid emitting under lock.
        self._emit_projection_changed()

    @property
    def files(self) -> list[FileRef]:
        """Current list of input file references (thread-safe copy)."""
        with QMutexLocker(self._mutex):
            return list(self._files)

    @property
    def has_files(self) -> bool:
        """Whether any input files are loaded."""
        with QMutexLocker(self._mutex):
            return len(self._files) > 0

    @property
    def status_message(self) -> str:
        """Current user-facing status message."""
        with QMutexLocker(self._mutex):
            return self._status_message

    @property
    def current_task_id(self) -> str | None:
        """Task ID of the currently active task, or None."""
        with QMutexLocker(self._mutex):
            return self._current_task_id

    @property
    def selected_file(self) -> FileRef | None:
        """The currently-selected file driving the right-panel projection.

        ``None`` when no file is selected (no-files state).  This is the
        single source of truth for right-panel routing; the MainWindow must
        not pick a panel from file suffixes or categories on its own.
        """
        with QMutexLocker(self._mutex):
            return self._selected_file

    @property
    def ui_projection(self) -> MainWindowUiProjection:
        """The current main-window UI projection (render-only).

        Derived via ``context → capabilities → projection`` from the current
        mode and selected file.  Read-only — callers must not mutate it.
        Returns a
        :class:`~docwen_gui.view_models.interaction.MainWindowUiProjection`.
        """
        return self._compute_projection()

    # ── Command methods (called by widgets → delegate to controller) ────

    def add_files(self, paths: list[str]) -> FileAddOutcome:
        """Add files to the current input list.

        Widgets call this when files are dropped, selected via dialog,
        or received via IPC.

        Args:
            paths: Absolute file paths to add.
        """
        from pathlib import Path

        from docwen_core.detection import (
            OOXML_SIGNATURE_INFO_METADATA_KEY,
        )
        from docwen_core.models import AdmissionDecision
        from docwen_core.models.file_ref import FileRef

        from .interaction import normalize_workflow_category

        new_refs: list[FileRef] = []
        rejected: list[tuple[str, str]] = []
        for p in paths:
            if not p:
                continue
            try:
                inspection = self._file_inspector(p)
                fmt = inspection.detected_format
                detected_category = inspection.workflow_category
                category = normalize_workflow_category(detected_category) or "other"

                warning = render_file_inspection_message(inspection)
                if inspection.decision is AdmissionDecision.BLOCK:
                    reason = render_file_inspection_message(inspection, prefer_reason=True).strip() or _t(
                        "components.file_drop.unsupported_type_msg",
                        "Unsupported file type: {filename}",
                        filename=Path(p).name,
                    )
                    rejected.append((p, reason))
                    continue

                metadata: dict[str, Any] = {
                    FILE_INSPECTION_METADATA_KEY: inspection.to_dict(),
                    OOXML_SIGNATURE_INFO_METADATA_KEY: dict(inspection.ooxml_signature),
                }

                ref = FileRef(
                    path=p,
                    format=fmt,
                    category=category,
                    warning_message=warning,
                    size_bytes=Path(p).stat().st_size if Path(p).is_file() else 0,
                    metadata=metadata,
                )
                new_refs.append(ref)
            except (ValueError, OSError):
                # Invalid/unreadable inputs never enter observable file state.
                rejected.append((p, _t("components.file_drop.file_unavailable", "File is unavailable")))

        if not new_refs:
            if rejected:
                self.set_status_message(rejected[0][1])
            return FileAddOutcome(added=(), rejected=tuple(rejected))

        files_snapshot: list[FileRef] | None = None
        added_refs: list[FileRef] = []
        file_count = 0
        with QMutexLocker(self._mutex):
            existing = {f.path for f in self._files}
            for ref in new_refs:
                if ref.path not in existing:
                    self._files.append(ref)
                    added_refs.append(ref)
                    existing.add(ref.path)

            if added_refs:
                files_snapshot = list(self._files)
            file_count = len(self._files)
        if files_snapshot is not None:
            self.files_changed.emit(files_snapshot)
            self._update_title_for_file_count(file_count)
        if rejected:
            self.set_status_message(rejected[0][1])
        return FileAddOutcome(added=tuple(added_refs), rejected=tuple(rejected))

    def remove_file(self, file_path: str) -> None:
        """Remove a single file from the input list.

        Args:
            file_path: The absolute path of the file to remove.
        """
        files_snapshot: list[FileRef] | None = None
        file_count = 0
        with QMutexLocker(self._mutex):
            before = len(self._files)
            self._files = [f for f in self._files if f.path != file_path]
            if len(self._files) != before:
                files_snapshot = list(self._files)
                file_count = len(self._files)
        if files_snapshot is not None:
            self.files_changed.emit(files_snapshot)
            self._update_title_for_file_count(file_count)

    def clear_files(self) -> None:
        """Remove all files from the input list."""
        with QMutexLocker(self._mutex):
            self._files.clear()
            self._selected_file = None
        self.files_cleared.emit()
        self._update_title_for_file_count(0)
        # Selection cleared → right panel must hide.  Emit outside the mutex.
        self._emit_projection_changed()

    def set_selected_file(self, file_ref: FileRef | None) -> None:
        """Set the currently-selected file driving the right-panel projection.

        This is the single entry point for selection from the main window
        (single-mode drop, batch list click, etc.).  Pass ``None`` (or use
        :meth:`clear_selected_file`) to clear.

        The projection is recomputed and ``ui_projection_changed`` is emitted
        only when the selection actually changes.
        """
        with QMutexLocker(self._mutex):
            if self._selected_file is file_ref:
                return
            if self._selected_file is not None and file_ref is not None and self._selected_file.path == file_ref.path:
                return
            self._selected_file = file_ref
        self._emit_projection_changed()

    def clear_selected_file(self) -> None:
        """Clear the selected-file source state and hide the right panel."""
        with QMutexLocker(self._mutex):
            if self._selected_file is None:
                return
            self._selected_file = None
        self._emit_projection_changed()

    def set_mode(self, mode: str) -> None:
        """Set single/batch mode (thread-safe)."""
        self.mode = mode  # property setter handles validation + signal + projection

    def set_status_message(self, message: str) -> None:
        """Update the status bar / info area message.

        Called by widgets and by TaskEventBridge on event dispatch.
        """
        changed = False
        with QMutexLocker(self._mutex):
            if self._status_message != message:
                self._status_message = message
                changed = True
        if changed:
            self.status_message_changed.emit(message)

    def request_shutdown(self) -> None:
        """Ask the owning window to begin graceful application shutdown."""
        self.shutdown_requested.emit()

    def begin_execution_telemetry(self, operation_id: str, runtime_task_ids: tuple[str, ...]) -> None:
        """Admit runtime telemetry identities for one GUI-owned execution."""
        normalized_operation_id = str(operation_id)
        accepted_ids = frozenset(str(task_id) for task_id in runtime_task_ids if str(task_id))
        if not normalized_operation_id or not accepted_ids:
            raise ValueError("Execution telemetry requires an operation id and at least one runtime task id")
        with QMutexLocker(self._mutex):
            self._active_execution_id = normalized_operation_id
            self._accepted_runtime_task_ids = accepted_ids
            self._ended_runtime_task_ids.clear()
            self._current_task_id = None

    def on_task_event(self, event_type: str, payload: dict[str, Any]) -> None:
        """Handle live runtime telemetry dispatched from TaskEventBridge.

        Widgets should NOT call this directly; the bridge does so
        on the main thread.

        Runtime terminal events close the live-task bookkeeping only.  They
        deliberately do not publish an operation result: the runtime event can
        reach the Qt event loop before the execution thread returns the full
        ``ConversionResult`` (artifacts, diagnostics, and aligned batch
        outcomes).  ``publish_execution_summary`` is the sole operation-level
        terminal publisher.

        Args:
            event_type: One of the task event type constants.
            payload: Event-specific data.
        """
        from docwen_core.events.task_events import (
            TASK_CANCELLED,
            TASK_COMPLETED,
            TASK_FAILED,
            TASK_PROGRESS,
            TASK_STARTED,
        )

        task_id = str(payload.get("task_id", ""))
        message = payload.get("message", "")

        publish_live_status = False
        with QMutexLocker(self._mutex):
            if task_id not in self._accepted_runtime_task_ids:
                return
            if event_type == TASK_STARTED and task_id not in self._ended_runtime_task_ids:
                self._current_task_id = task_id
                publish_live_status = True
            elif event_type == TASK_PROGRESS and task_id not in self._ended_runtime_task_ids:
                publish_live_status = True
            elif event_type in (TASK_COMPLETED, TASK_FAILED, TASK_CANCELLED):
                self._ended_runtime_task_ids.add(task_id)
                if self._current_task_id == task_id:
                    self._current_task_id = None

        if event_type == TASK_STARTED and publish_live_status:
            self.set_status_message(_t("main_window.task_processing_status", message=message))
        elif event_type == TASK_PROGRESS and publish_live_status:
            pct = payload.get("percent", 0.0)
            self.set_status_message(_t("main_window.task_progress_status", percent=f"{pct:.0f}", message=message))

    def publish_execution_summary(self, status: str, payload: dict[str, Any] | None = None) -> None:
        """Publish one authoritative operation terminal after result projection.

        The caller must first commit batch rows, history, and the detailed task
        summary.  This method then closes telemetry state and exposes exactly
        one operation-level terminal signal to external GUI observers.
        """
        terminal_payload = dict(payload or {})
        terminal_payload["status"] = status
        message = str(terminal_payload.get("message", ""))
        operation_id = str(terminal_payload.get("task_id", ""))

        with QMutexLocker(self._mutex):
            if not self._active_execution_id or self._active_execution_id == operation_id:
                self._active_execution_id = None
                self._accepted_runtime_task_ids = frozenset()
                self._ended_runtime_task_ids.clear()
                self._current_task_id = None

        if status in {STATUS_COMPLETED, "partial"}:
            self.set_status_message(_t("main_window.task_completed_status"))
        elif status == STATUS_FAILED:
            self.set_status_message(_t("main_window.task_failed_status", message=message))
        elif status == STATUS_CANCELLED:
            self.set_status_message(_t("main_window.task_cancelled_status"))
        else:
            raise ValueError(f"Unsupported execution terminal status: {status!r}")
        self.task_summary_changed.emit(terminal_payload)

    # ── IPC command handling ─────────────────────────────────────────────

    def handle_ipc_command(self, action: str, file_path: str | None = None) -> None:
        """Handle a command received from another process via IPC.

        Called by the main window when runtime/control delivers a command.

        Args:
            action: The command action (``"activate"``, ``"add_file"``,
                ``"open_file"``).
            file_path: Optional file path associated with the command.
        """
        if action in ("add_file", "open_file") and file_path:
            path = Path(file_path)
            if not path.exists():
                logger.warning("IPC file does not exist — skipping: %s", file_path)
                self.set_status_message(_t("info_area.ipc_file_missing", "IPC file not found: {path}", path=file_path))
                return
            outcome = self.add_files([str(path)])
            if outcome.added:
                self.ipc_file_received.emit(str(path))
            # Always bring window to front when files are added.
            self.window_activation_requested.emit()
        elif action == "activate":
            self.window_activation_requested.emit()
        else:
            logger.debug("Unknown or no-op IPC action: %r", action)

    def request_window_activation(self) -> None:
        """Request that the main window be activated (brought to front).

        Emits ``window_activation_requested`` which the MainWindow connects
        to ``bring_to_front()``.

        Kept intentionally as VM public API: this is the Qt-slot/semantic
        wrapper around ``window_activation_requested.emit()``. It has no
        static production caller today, but it documents the activation
        request as an addressable VM method and remains available for
        dynamic dispatch (e.g. IPC ``activate`` wiring, future shortcuts).
        Removing it would drop the explicit activation entry point on the
        VM surface.
        """
        self.window_activation_requested.emit()

    # ── Internal helpers ────────────────────────────────────────────────

    def _compute_projection(self) -> MainWindowUiProjection:
        """Build the current UI projection from mode + selected file.

        Pure delegation to :mod:`docwen_gui.view_models.interaction` — the
        routing rules live there, Qt-free and unit-tested.  The VM only
        supplies the source state.
        """
        from .interaction import UiMode, build_interaction_context, project_main_window_ui

        with QMutexLocker(self._mutex):
            mode = self._mode
            selected = self._selected_file
        ui_mode = UiMode(mode)
        category = getattr(selected, "category", None) if selected is not None else None
        fmt = getattr(selected, "format", None) if selected is not None else None
        path = getattr(selected, "path", None) if selected is not None else None
        context = build_interaction_context(
            mode=ui_mode,
            category=category,
            current_format=fmt,
            file_path=path,
            selected_file=selected,
        )
        return project_main_window_ui(context)

    def _emit_projection_changed(self) -> None:
        """Recompute the projection and emit ``ui_projection_changed``."""
        self.ui_projection_changed.emit(self._compute_projection())

    def _update_title_from_files(self) -> None:
        """Update window title to include file count."""
        count = len(self._files)
        self._update_title_for_file_count(count)

    def _update_title_for_file_count(self, count: int) -> None:
        """Update window title to include a known file count."""
        try:
            from docwen_gui.i18n import t as _i18n_t

            base = _i18n_t("main_window.window_title", "DocWen Offline")
        except ImportError:
            base = "DocWen Offline"
        title = base if count == 0 else f"{base} [{count} file{'s' if count != 1 else ''}]"
        self._title = title
        self.title_changed.emit(title)


__all__ = [
    "DEFAULT_MODE",
    "STATUS_COMPLETED",
    "STATUS_FAILED",
    "STATUS_IDLE",
    "STATUS_PROCESSING",
    "FileAddOutcome",
    "MainWindowViewModel",
]
