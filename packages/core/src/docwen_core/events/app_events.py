"""Application-level event types.

These events flow between application, runtime, and presenters
(GUI / CLI).  They are *not* the same as ``TaskEvent`` — application
events describe user-level state changes (batch started, file added,
files dropped, etc.).

Some event types are inherently GUI-specific and will never fire in
a CLI-only or headless deployment.  See individual constant docstrings.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import UTC, datetime
from typing import Any

# ── Event type constants ─────────────────────────────────────────────
#
# Convention:
#   Generic events     — fire in both GUI and CLI modes
#   GUI-only events    — fire only when a GUI window is present
#
# The string constants are defined here (in core) so that both
# application and runtime can reference them without a GUI dependency.
# The *producers* of GUI-only events live in the GUI layer.

# Generic
APP_STARTED = "app_started"
BATCH_STARTED = "batch_started"
BATCH_COMPLETED = "batch_completed"
BATCH_CANCELLED = "batch_cancelled"
TASK_ENQUEUED = "task_enqueued"
TASK_FINISHED = "task_finished"
FILES_ADDED = "files_added"
CONFIG_CHANGED = "config_changed"

# GUI-only (produced by the GUI layer, never by CLI)
FILES_DROPPED = "files_dropped"  # drag-and-drop onto the GUI window
IPC_FILE_RECEIVED = "ipc_file_received"  # second instance launched with file args
WINDOW_ACTIVATED = "window_activated"  # existing window brought to front


@dataclass(slots=True)
class AppEvent:
    """A user-level application event."""

    event_type: str
    """Discriminator (see constants above)."""

    timestamp: str = ""
    """ISO-8601 UTC timestamp.  Auto-filled if empty."""

    payload: dict[str, Any] = field(default_factory=dict)
    """Event-specific data."""

    def __post_init__(self) -> None:
        if not self.timestamp:
            self.timestamp = datetime.now(UTC).isoformat()

    def to_dict(self) -> dict[str, Any]:
        return {
            "event_type": self.event_type,
            "timestamp": self.timestamp,
            "payload": dict(self.payload),
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> AppEvent:
        return cls(
            event_type=data["event_type"],
            timestamp=data.get("timestamp", ""),
            payload=dict(data.get("payload", {})),
        )
