"""InfoAreaViewModel — observable state for the InfoArea widget.

This is the single source of truth for the InfoArea's observable state.
Widgets bind to its signals and properties; user actions flow through
method calls.

State managed:
- History messages: 100 max, 250ms dedup (six-tuple signature), HH:MM:SS timestamps
- Transient messages: priority (error>terminal>progress>processing), TTL (3000ms default, 4000ms terminal), progress throttle (250ms)
- Task summary: drives guide button rendering and status display
- Activity animation: 300ms interval dot cycle
- Guide actions: success/partial/failed/cancelled button combinations
"""

from __future__ import annotations

import logging
import time
from dataclasses import dataclass, replace
from datetime import datetime

from PySide6.QtCore import QObject, QTimer, Signal

from docwen_gui.i18n import t as _t

logger = logging.getLogger(__name__)

# ── Transient priority (lower = higher priority) ──────────────────────────
_TRANSIENT_PRIORITY: dict[str, int] = {
    "error": 0,
    "terminal": 1,
    "progress": 2,
    "processing": 3,
}

# ── Throttle windows for transient types (ms) ─────────────────────────────
_TRANSIENT_THROTTLE_MS: dict[str, int] = {
    "progress": 250,
}

# ── Message types considered "important" (preserved on overflow) ─────────
_IMPORTANT_TYPES: frozenset[str] = frozenset({"success", "danger", "warning"})

# ── Task guide action label i18n keys ─────────────────────────────────────
_TASK_GUIDE_LABELS: dict[str, str] = {
    "open_output_dir": "info_area.task_guide_open_output_dir",
    "view_failed_details": "info_area.task_guide_view_failed_details",
    "retry_failed": "info_area.task_guide_retry_failed",
    "add_more_files": "info_area.task_guide_add_more_files",
}

# ── Task states that trigger guide row ────────────────────────────────────
_GUIDE_ELIGIBLE_STATES: frozenset[str] = frozenset(
    {
        "success",
        "partial",
        "failed",
        "cancelled",
    }
)

# ``_refresh_status`` builds ``info_area.task_state_{state}`` dynamically.
# Keep that lookup finite and fail closed before any unknown state can enter
# the view-model.  The empty state is reserved for the idle DTO.
TASK_SUMMARY_STATES = frozenset({"active", "success", "partial", "failed", "cancelled"})
TASK_SUMMARY_TONES = frozenset({"info", "success", "warning", "danger"})


# ── Data classes ──────────────────────────────────────────────────────────


@dataclass(slots=True)
class _PendingTransient:
    """A transient update queued behind the throttle window."""

    message: str
    message_type: str
    priority: int
    ttl_ms: int
    source: str
    version: int


@dataclass(slots=True)
class HistoryRowData:
    """Normalised data for a single history row (widget-independent)."""

    timestamp: str
    message: str
    message_type: str
    show_location: bool = False
    file_path: str = ""
    navigate_file_path: str = ""
    operation_id: str = ""


@dataclass
class TaskSummaryState:
    """Aggregated task summary state for cleaner batch summary display.

    Replaces the previous scalar ``_task_summary_*`` fields in
    ``InfoAreaViewModel`` with a single DTO.  All count fields
    default to zero.
    """

    state: str = ""  # active / success / partial / failed / cancelled
    tone: str = "info"  # info / success / warning / danger
    operation_id: str = ""
    current_file: str = ""
    completed_count: int = 0
    total_count: int = 0
    failed_count: int = 0
    skipped_count: int = 0
    cancelled_count: int = 0
    navigate_path: str = ""
    navigation_kind: str = ""

    def __post_init__(self) -> None:
        state = str(self.state).strip().lower()
        tone = str(self.tone).strip().lower()
        if state and state not in TASK_SUMMARY_STATES:
            raise ValueError(f"state must be empty or one of {sorted(TASK_SUMMARY_STATES)!r}")
        if tone not in TASK_SUMMARY_TONES:
            raise ValueError(f"tone must be one of {sorted(TASK_SUMMARY_TONES)!r}")
        self.state = state
        self.tone = tone


# ── ViewModel ─────────────────────────────────────────────────────────────


class InfoAreaViewModel(QObject):
    """State source of truth for the InfoArea widget.

    Manages history messages, transient messages, task summary,
    guide actions, activity animation, and status display state.

    Signals:
        state_changed: Emitted whenever display-relevant state changes.
        history_navigation_requested: Emitted when user requests navigation
            from a history row or task summary.
        task_guide_action_requested: Emitted when user clicks a guide button.
    """

    # ── Signals ──────────────────────────────────────────────────────────

    state_changed = Signal()
    """Emitted when any display-relevant state changes — widgets rebind."""

    history_navigation_requested = Signal(str)
    """Emitted when user clicks an interactive status summary (target path)."""

    task_guide_action_requested = Signal(str, str)
    """Emitted when user clicks a guide button: (action_key, target_path)."""

    location_requested = Signal(str)
    """Emitted when user clicks a location button: (file_path)."""

    # ── Construction ──────────────────────────────────────────────────────

    def __init__(self, parent: QObject | None = None) -> None:
        super().__init__(parent)
        self.max_messages: int = 100

        # ── History state ────────────────────────────────────────────────
        self._history_rows: list[HistoryRowData] = []
        self._last_message_signature: tuple | None = None
        self._last_message_time: float = 0.0

        # ── Transient state ──────────────────────────────────────────────
        # _transient_messages: key -> (message, message_type, priority)
        self._transient_messages: dict[str, tuple[str, str, int]] = {}
        self._transient_timers: dict[str, QTimer] = {}
        self._transient_versions: dict[str, int] = {}
        self._transient_sources: dict[str, str] = {}
        self._transient_last_displayed_at: dict[str, float] = {}
        self._pending_transient_updates: dict[str, _PendingTransient] = {}
        self._transient_flush_timers: dict[str, QTimer] = {}
        # Invalidates already-queued timer callbacks when the current working
        # session is reset.  Stopping a QTimer prevents future timeouts, but a
        # callback that was already queued in Qt's event loop must also be made
        # harmless before the same transient key can be reused.
        self._transient_generation: int = 0

        # ── Task summary (DTO) ─────────────────────────────────────────
        self._task_summary: TaskSummaryState = TaskSummaryState()
        self._guide_actions: list[dict[str, str]] = []

        # ── Activity animation ───────────────────────────────────────────
        self._activity_timer: QTimer | None = None
        self._activity_step: int = 0
        self._activity_base_meta: str = ""
        self._activity_enabled: bool = False

        # ── Computed display state ───────────────────────────────────────
        self._status_meta_text: str = ""
        self._status_summary_text: str = ""
        self._status_source: str = "idle"
        self._status_tone: str = "secondary"
        self._status_action_target: str = ""
        self._guide_visible: bool = False

        # Initial refresh
        self._refresh_status()

    # ── Public properties (read-only display state) ──────────────────────

    @property
    def status_meta_text(self) -> str:
        """Current status meta text (e.g. overview source + activity dots)."""
        return self._status_meta_text

    @property
    def status_summary_text(self) -> str:
        """Current status summary text."""
        return self._status_summary_text

    @property
    def status_source(self) -> str:
        """Current status source: idle / history / transient / task."""
        return self._status_source

    @property
    def status_tone(self) -> str:
        """Current status tone: info / success / warning / danger / secondary."""
        return self._status_tone

    @property
    def status_action_target(self) -> str:
        """Current interactive navigation target path (or empty)."""
        return self._status_action_target

    @property
    def guide_visible(self) -> bool:
        """Whether the task guide row should be visible."""
        return self._guide_visible

    @property
    def guide_actions(self) -> list[dict[str, str]]:
        """Current guide action list: [{"action_key": ..., "target_path": ...}, ...]."""
        return list(self._guide_actions)

    @property
    def history_rows(self) -> list[HistoryRowData]:
        """Current history rows (copy for rendering)."""
        return list(self._history_rows)

    @property
    def message_count(self) -> int:
        """Number of history messages."""
        return len(self._history_rows)

    @property
    def message_types(self) -> list[str]:
        """List of message types in order (for testing)."""
        return [row.message_type for row in self._history_rows]

    @property
    def transient_count(self) -> int:
        """Number of active transient messages."""
        return len(self._transient_messages)

    @property
    def has_task_summary(self) -> bool:
        """Whether a task summary is currently set."""
        return bool(self._task_summary.state)

    @property
    def task_summary(self) -> TaskSummaryState:
        """Current task summary snapshot for notification/status consumers."""
        return replace(self._task_summary)

    @property
    def activity_enabled(self) -> bool:
        """Whether the activity animation is active."""
        return self._activity_enabled

    @property
    def activity_meta_text(self) -> str:
        """The current activity meta text (base + dots)."""
        return self._format_activity_meta(self._activity_base_meta, self._activity_step)

    # ── Public methods: history ───────────────────────────────────────────

    def add_message(
        self,
        message: str,
        message_type: str = "secondary",
        show_location: bool = False,
        file_path: str | None = None,
        navigate_file_path: str | None = None,
        operation_id: str | None = None,
    ) -> None:
        """Add a message to the history area.

        Deduplicates against the previous message using a six-tuple
        signature within a 250ms window.

        Args:
            message: The user-visible message text.
            message_type: Semantic type (info / success / warning / danger / secondary).
            show_location: Whether to show a location button for this message.
            file_path: File path for the location button.
            navigate_file_path: Path to navigate to when the status summary is clicked.
            operation_id: Associated operation ID (triggers terminal transient).
        """
        navigation_target = navigate_file_path or ""
        resolved_operation_id = operation_id or ""
        signature = (
            message,
            message_type,
            bool(show_location),
            file_path or "",
            navigation_target,
            resolved_operation_id,
        )
        now = time.monotonic()

        # 250ms dedup window
        if signature == self._last_message_signature and (now - self._last_message_time) < 0.25:
            return

        self._last_message_signature = signature
        self._last_message_time = now

        timestamp = datetime.now().strftime("%H:%M:%S")

        row = HistoryRowData(
            timestamp=timestamp,
            message=message,
            message_type=message_type,
            show_location=bool(show_location),
            file_path=file_path or "",
            navigate_file_path=navigation_target,
            operation_id=resolved_operation_id,
        )
        self._history_rows.append(row)
        self._enforce_message_limit()

        # Terminal transient for success/danger/warning with operation_id
        if resolved_operation_id and message_type in _IMPORTANT_TYPES:
            self._publish_terminal_transient(resolved_operation_id, message, message_type=message_type)

        self._refresh_status()
        logger.debug("History message: [%s] %s", message_type, message)

    def remove_history_row_at(self, index: int) -> None:
        """Remove a history row by index."""
        if 0 <= index < len(self._history_rows):
            self._history_rows.pop(index)
            self._refresh_status()

    def clear_history(self) -> None:
        """Clear all history rows."""
        self._history_rows.clear()
        self._last_message_signature = None
        self._last_message_time = 0.0
        self._refresh_status()

    def reset_session(self) -> None:
        """Atomically reset the information state for the current work session.

        This intentionally clears only in-memory, session-scoped UI state:
        history rows, transient messages (including pending updates and timers),
        task summary/guide actions, and activity animation.  Persistent logs,
        output files, recent files, and settings are owned elsewhere and remain
        untouched.
        """
        self._history_rows.clear()
        self._last_message_signature = None
        self._last_message_time = 0.0

        self._transient_generation += 1
        self._clear_transient_state()

        self._task_summary = TaskSummaryState()
        self._guide_actions.clear()

        self._activity_enabled = False
        self._activity_step = 0
        self._activity_base_meta = ""
        if self._activity_timer is not None and self._activity_timer.isActive():
            self._activity_timer.stop()

        # One refresh keeps the reset atomic from the widget's perspective.
        self._refresh_status()

    # ── Public methods: transient ─────────────────────────────────────────

    @staticmethod
    def _transient_key_base(key: str) -> str:
        """Extract the base type from a transient key (e.g. 'progress:op-123' -> 'progress')."""
        return key.split(":", 1)[0]

    def set_transient_message(
        self,
        key: str,
        message: str,
        message_type: str = "secondary",
        ttl_ms: int = 3000,
        *,
        source: str | None = None,
        force_refresh: bool = False,
    ) -> None:
        """Set or update a transient message.

        Transient messages have priority-based display (lower number = higher priority)
        and auto-expire after ``ttl_ms``.

        Args:
            key: Unique key for this transient (e.g. "progress:op-1234").
            message: User-visible message text.
            message_type: Semantic type (info / success / warning / danger / secondary).
            ttl_ms: Time-to-live in milliseconds (0 = no auto-expiry).
            source: Source identifier for grouping (defaults to key).
            force_refresh: If True, bypass throttle and display immediately.
        """
        key_base = self._transient_key_base(key)
        priority = _TRANSIENT_PRIORITY.get(key_base, 99)
        resolved_source = source or key
        version = self._transient_versions.get(key, 0) + 1
        self._transient_versions[key] = version

        payload = _PendingTransient(
            message=message,
            message_type=message_type,
            priority=priority,
            ttl_ms=ttl_ms,
            source=resolved_source,
            version=version,
        )

        throttle_ms = _TRANSIENT_THROTTLE_MS.get(key_base, 0)
        last_displayed_at = self._transient_last_displayed_at.get(key, 0.0)
        now = time.monotonic()

        should_display_now = (
            force_refresh
            or throttle_ms <= 0
            or key not in self._transient_messages
            or (now - last_displayed_at) >= (throttle_ms / 1000.0)
        )

        if should_display_now:
            self._apply_transient_update(key, payload)
        else:
            self._pending_transient_updates[key] = payload
            self._restart_transient_ttl_timer(key, version, ttl_ms)
            delay_ms = max(0, throttle_ms - int((now - last_displayed_at) * 1000))
            self._schedule_transient_flush(key, version, delay_ms)

    def clear_transient_message(self, key: str) -> None:
        """Clear a transient message by key."""
        self._transient_messages.pop(key, None)
        self._transient_versions.pop(key, None)
        self._transient_sources.pop(key, None)
        self._transient_last_displayed_at.pop(key, None)
        self._pending_transient_updates.pop(key, None)

        if key in self._transient_timers:
            self._transient_timers[key].stop()
            del self._transient_timers[key]
        if key in self._transient_flush_timers:
            self._transient_flush_timers[key].stop()
            del self._transient_flush_timers[key]

        self._refresh_status()

    def clear_all_transients(self) -> None:
        """Clear all transient messages and stop all related timers."""
        self._transient_generation += 1
        self._clear_transient_state()
        self._refresh_status()

    def _clear_transient_state(self) -> None:
        """Clear transient state without emitting a display refresh."""
        self._transient_messages.clear()
        self._transient_versions.clear()
        self._transient_sources.clear()
        self._transient_last_displayed_at.clear()
        self._pending_transient_updates.clear()

        for timer in self._transient_timers.values():
            timer.stop()
        self._transient_timers.clear()
        for timer in self._transient_flush_timers.values():
            timer.stop()
        self._transient_flush_timers.clear()

    # ── Public methods: task summary ──────────────────────────────────────

    def set_task_summary(
        self,
        *,
        operation_id: str = "",
        current_file: str = "",
        current_file_path: str = "",
        completed_count: int = 0,
        total_count: int = 0,
        failed_count: int = 0,
        skipped_count: int = 0,
        cancelled_count: int = 0,
        state: str = "active",
        tone: str = "info",
        navigate_file_path: str = "",
        navigation_kind: str = "",
        guide_actions: list[dict[str, str]] | None = None,
    ) -> None:
        """Set or update the task summary state.

        Args:
            operation_id: Unique operation identifier.
            current_file: Name of the currently processing file.
            current_file_path: Absolute path of the current file.
            completed_count: Number of completed items.
            total_count: Total number of items.
            failed_count: Number of failed items.
            skipped_count: Number of skipped items.
            cancelled_count: Number of cancelled items.
            state: Task state: active / success / partial / failed / cancelled.
            tone: Semantic tone: info / success / warning / danger.
            navigate_file_path: Path for status summary click navigation.
            navigation_kind: Navigation kind hint (e.g. "current", "failed").
            guide_actions: List of guide action dicts for the completed state.
        """
        self._task_summary = TaskSummaryState(
            state=state,
            tone=tone,
            operation_id=operation_id,
            current_file=current_file,
            completed_count=completed_count,
            total_count=total_count,
            failed_count=failed_count,
            skipped_count=skipped_count,
            cancelled_count=cancelled_count,
            navigate_path=navigate_file_path,
            navigation_kind=navigation_kind,
        )
        self._guide_actions = list(guide_actions or [])
        self._refresh_status()

    def clear_task_summary(self) -> None:
        """Clear the task summary and return to idle/history display."""
        self._task_summary = TaskSummaryState()
        self._guide_actions.clear()
        self._refresh_status()

    # ── Public methods: activity animation ───────────────────────────────

    def start_activity_animation(self, base_meta: str) -> None:
        """Start the activity animation with the given base meta text."""
        self._activity_base_meta = base_meta
        self._activity_enabled = True
        self._activity_step = 0
        self._ensure_activity_timer()
        if self._activity_timer is not None and not self._activity_timer.isActive():
            self._activity_timer.start()
        self.state_changed.emit()

    def stop_activity_animation(self) -> None:
        """Stop the activity animation."""
        self._activity_enabled = False
        if self._activity_timer is not None and self._activity_timer.isActive():
            self._activity_timer.stop()
        self.state_changed.emit()

    def stop_all_timers(self) -> None:
        """Stop all timers (transient TTL, flush, activity). Call before destruction."""
        for timer in self._transient_timers.values():
            timer.stop()
        self._transient_timers.clear()
        for timer in self._transient_flush_timers.values():
            timer.stop()
        self._transient_flush_timers.clear()
        if self._activity_timer is not None:
            self._activity_timer.stop()
            self._activity_timer = None

    # ── Public methods: actions ──────────────────────────────────────────

    def request_navigation(self) -> None:
        """Request navigation to the current status action target."""
        if self._status_action_target:
            self.history_navigation_requested.emit(self._status_action_target)

    def request_guide_action(self, action_key: str, target_path: str = "") -> None:
        """Request a guide action (called by widget on button click)."""
        self.task_guide_action_requested.emit(action_key, target_path)

    def request_location(self, file_path: str) -> None:
        """Request opening the file location for a specific path."""
        if file_path:
            self.location_requested.emit(file_path)

    # ── Internal: message limit enforcement ──────────────────────────────

    def _enforce_message_limit(self) -> None:
        """Remove oldest non-important entries when over the message limit.

        Important types (success, danger, warning) are preserved preferentially.
        If ALL entries are important, the first (oldest) is removed.
        """
        while len(self._history_rows) > self.max_messages:
            removal_index = 0
            for idx, row in enumerate(self._history_rows):
                if row.message_type not in _IMPORTANT_TYPES:
                    removal_index = idx
                    break
            self._history_rows.pop(removal_index)

    # ── Internal: transient management ────────────────────────────────────

    def _apply_transient_update(self, key: str, payload: _PendingTransient) -> None:
        """Apply a pending transient update, checking version for staleness."""
        current_version = self._transient_versions.get(key)
        if current_version != payload.version:
            return

        self._pending_transient_updates.pop(key, None)
        self._transient_messages[key] = (
            payload.message,
            payload.message_type,
            payload.priority,
        )
        self._transient_sources[key] = payload.source
        self._transient_last_displayed_at[key] = time.monotonic()

        if key in self._transient_flush_timers:
            self._transient_flush_timers[key].stop()
            del self._transient_flush_timers[key]

        self._restart_transient_ttl_timer(key, payload.version, payload.ttl_ms)
        self._refresh_status()

    def _restart_transient_ttl_timer(self, key: str, version: int, ttl_ms: int) -> None:
        """(Re)start the auto-expiry timer for a transient message."""
        if key in self._transient_timers:
            self._transient_timers[key].stop()
            del self._transient_timers[key]

        if ttl_ms <= 0:
            return

        timer = QTimer(self)
        timer.setSingleShot(True)
        generation = self._transient_generation
        timer.timeout.connect(lambda k=key, v=version, g=generation: self._expire_transient(k, v, g))
        timer.start(ttl_ms)
        self._transient_timers[key] = timer

    def _schedule_transient_flush(self, key: str, version: int, delay_ms: int) -> None:
        """Schedule a delayed flush for a throttled transient update."""
        if key in self._transient_flush_timers:
            self._transient_flush_timers[key].stop()
            del self._transient_flush_timers[key]

        timer = QTimer(self)
        timer.setSingleShot(True)
        generation = self._transient_generation
        timer.timeout.connect(lambda k=key, v=version, g=generation: self._flush_pending_transient(k, v, g))
        timer.start(delay_ms)
        self._transient_flush_timers[key] = timer

    def _flush_pending_transient(self, key: str, version: int, generation: int | None = None) -> None:
        """Flush a pending transient update (called by delayed timer)."""
        if generation is not None and generation != self._transient_generation:
            return
        payload = self._pending_transient_updates.get(key)
        if payload is None or payload.version != version:
            return
        self._apply_transient_update(key, payload)

    def _expire_transient(self, key: str, version: int, generation: int | None = None) -> None:
        """Handle transient TTL expiry (called by single-shot timer)."""
        if generation is not None and generation != self._transient_generation:
            return
        if self._transient_versions.get(key) != version:
            return
        self.clear_transient_message(key)

    def _clear_transient_source(self, source: str) -> None:
        """Clear all transient messages from a given source."""
        if not source:
            return
        keys_to_clear = {k for k, s in self._transient_sources.items() if s == source}
        keys_to_clear.update(k for k, p in self._pending_transient_updates.items() if p.source == source)
        for key in keys_to_clear:
            self.clear_transient_message(key)

    def _publish_terminal_transient(self, source: str, message: str, *, message_type: str) -> None:
        """Publish a terminal transient (brief result flash after task end).

        Clears all existing transients from the same source, then shows
        a terminal transient with TTL=4000ms and force_refresh=True.
        """
        self._clear_transient_source(source)
        self.set_transient_message(
            f"terminal:{source}",
            message,
            message_type=message_type,
            ttl_ms=4000,
            source=source,
            force_refresh=True,
        )

    # ── Internal: status refresh ──────────────────────────────────────────

    def _refresh_status(self) -> None:
        """Recalculate display state from current data sources.

        Priority: transient > task_summary > history > idle.
        """
        activity_enabled = False
        overview_source = "idle"
        badge_tone = "secondary"
        overview_meta = ""
        action_target = ""

        # Determine best transient
        transient_message: str | None = None
        transient_theme = "secondary"
        if self._transient_messages:
            # Lowest priority number = highest display priority
            best_key, best = min(
                self._transient_messages.items(),
                key=lambda item: item[1][2],
            )
            transient_message, transient_theme, _ = best

        ts = self._task_summary
        has_task = bool(ts.state)

        if has_task:
            state = ts.state
            current_file = ts.current_file or _t("common.ready", "Ready")
            overview_meta = _t(
                f"info_area.task_state_{state}",
                state.replace("_", " ").title(),
            )
            completed = ts.completed_count
            total = ts.total_count
            failed = ts.failed_count
            op_id = ts.operation_id
            progress_lines = [
                _t(
                    "info_area.task_progress_detail",
                    "Completed: {completed}/{total}, Failed: {failed} [{operation_id}]",
                    completed=completed,
                    total=total,
                    failed=failed,
                    operation_id=op_id,
                )
            ]
            if ts.skipped_count:
                progress_lines.append(
                    _t(
                        "info_area.task_skipped_count",
                        "Skipped: {skipped}",
                        skipped=ts.skipped_count,
                    )
                )
            if ts.cancelled_count:
                progress_lines.append(
                    _t(
                        "info_area.task_cancelled_count",
                        "Cancelled: {cancelled}",
                        cancelled=ts.cancelled_count,
                    )
                )
            summary_message = "\n".join(
                [
                    _t("info_area.task_current_file", "Current: {name}", name=current_file),
                    *progress_lines,
                ]
            )

            message = transient_message or summary_message
            badge_tone = transient_theme if transient_message else ts.tone
            overview_source = "transient" if transient_message else "task"
            action_target = ts.navigate_path
            activity_enabled = state == "active" and transient_message is None
        else:
            if transient_message is not None:
                message = transient_message
                badge_tone = transient_theme
                overview_meta = _t(
                    "info_area.transient_meta",
                    "Processing ({count})",
                    count=len(self._transient_messages),
                )
                overview_source = "transient"
                # Activity animation for progress/processing transients
                best_key_base = ""
                if self._transient_messages:
                    best_key, _ = min(
                        self._transient_messages.items(),
                        key=lambda item: item[1][2],
                    )
                    best_key_base = self._transient_key_base(best_key)
                activity_enabled = best_key_base in {"progress", "processing"}
            elif self._history_rows:
                latest = self._history_rows[-1]
                message = latest.message
                badge_tone = latest.message_type
                overview_meta = _t(
                    "info_area.history_meta",
                    "History ({count})",
                    count=len(self._history_rows),
                )
                overview_source = "history"
                action_target = latest.navigate_file_path
            else:
                message = _t("common.ready", "Ready")
                overview_meta = _t("info_area.history_meta", "History (0)", count=0)

        # Update state
        self._status_meta_text = overview_meta
        self._status_summary_text = message
        self._status_source = overview_source
        self._status_tone = badge_tone
        self._status_action_target = action_target

        # Manage activity animation
        if activity_enabled:
            if not self._activity_enabled:
                self._activity_base_meta = overview_meta
                self._activity_enabled = True
                self._activity_step = 0
                self._ensure_activity_timer()
                if self._activity_timer is not None and not self._activity_timer.isActive():
                    self._activity_timer.start()
        else:
            if self._activity_enabled:
                self._activity_enabled = False
                if self._activity_timer is not None and self._activity_timer.isActive():
                    self._activity_timer.stop()

        # Refresh guide buttons
        self._refresh_guide()

        # Emit
        self.state_changed.emit()

    # ── Internal: activity animation ──────────────────────────────────────

    def _ensure_activity_timer(self) -> None:
        """Create the activity animation timer if it doesn't exist."""
        if self._activity_timer is not None:
            return
        timer = QTimer(self)
        timer.setInterval(300)
        timer.timeout.connect(self._tick_activity_animation)
        self._activity_timer = timer

    @staticmethod
    def _format_activity_meta(base: str, step: int) -> str:
        """Format activity meta text with cycling dots."""
        dots = step % 4
        if dots <= 0:
            return base
        return f"{base}{'.' * dots}"

    def _tick_activity_animation(self) -> None:
        """Advance the activity animation by one step."""
        if not self._activity_enabled:
            if self._activity_timer is not None and self._activity_timer.isActive():
                self._activity_timer.stop()
            return
        self._activity_step += 1
        # Update the meta text with animated dots
        self._status_meta_text = self._format_activity_meta(self._activity_base_meta, self._activity_step)
        self.state_changed.emit()

    # ── Internal: guide buttons ───────────────────────────────────────────

    def _refresh_guide(self) -> None:
        """Determine which guide buttons to show based on task summary state."""
        state = self._task_summary.state
        guide_actions = self._guide_actions

        if not state or state not in _GUIDE_ELIGIBLE_STATES:
            self._guide_visible = False
            self._guide_actions.clear()
            return

        if not guide_actions:
            self._guide_visible = False
            self._guide_actions.clear()
            return

        self._guide_actions = list(guide_actions)
        self._guide_visible = True

    @staticmethod
    def compute_guide_actions(
        state: str,
        output_dir: str = "",
        failed_details_path: str = "",
        retry_available: bool = False,
    ) -> list[dict[str, str]]:
        """Compute the guide action list based on task completion state.

        This static method encapsulates the guide button combination rules:
        - All success: open_output_dir + add_more_files
        - Has failures: open_output_dir + view_failed_details + retry_failed
        - Cancelled: only add_more_files

        Args:
            state: Task state (success / partial / failed / cancelled).
            output_dir: Path to the output directory.
            failed_details_path: Path to failed item details.
            retry_available: Whether retry is available.

        Returns:
            List of guide action dicts with action_key and target_path.
        """
        actions: list[dict[str, str]] = []

        if state == "cancelled":
            actions.append({"action_key": "add_more_files", "target_path": ""})
            return actions

        # For success, partial, and failed states
        if output_dir:
            actions.append({"action_key": "open_output_dir", "target_path": output_dir})
        else:
            actions.append({"action_key": "open_output_dir", "target_path": ""})

        if state in ("partial", "failed"):
            if failed_details_path:
                actions.append(
                    {
                        "action_key": "view_failed_details",
                        "target_path": failed_details_path,
                    }
                )
            else:
                actions.append(
                    {
                        "action_key": "view_failed_details",
                        "target_path": "",
                    }
                )
            if retry_available:
                actions.append({"action_key": "retry_failed", "target_path": ""})

        actions.append({"action_key": "add_more_files", "target_path": ""})
        return actions

    @staticmethod
    def guide_action_label(action_key: str) -> str:
        """Get the translated label for a guide action key."""
        i18n_key = _TASK_GUIDE_LABELS.get(action_key, "common.ok")
        return _t(i18n_key, action_key.replace("_", " ").title())


__all__ = [
    "_IMPORTANT_TYPES",
    "HistoryRowData",
    "InfoAreaViewModel",
]
