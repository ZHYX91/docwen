"""TaskEvent — events emitted during task execution."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import UTC, datetime
from typing import Any


@dataclass(slots=True)
class TaskEvent:
    """An event emitted during a task's lifecycle.

    Task events flow from worker → runtime → application → presenter.
    They are the only mechanism through which a plugin communicates
    progress or status changes.
    """

    task_id: str
    """The task that emitted this event."""

    event_type: str
    """Event type discriminator.

    Standard types:
    - ``"task_started"``
    - ``"task_progress"``
    - ``"task_completed"``
    - ``"task_failed"``
    - ``"task_cancelled"``
    - ``"artifact_ready"``
    - ``"diagnostic"``
    """

    sequence: int
    """Monotonically increasing sequence number for ordering."""

    timestamp: str = ""
    """ISO-8601 UTC timestamp.  Auto-filled if empty."""

    payload: dict[str, Any] = field(default_factory=dict)
    """Type-specific payload (progress percentage, message, etc.)."""

    def __post_init__(self) -> None:
        if not self.timestamp:
            self.timestamp = datetime.now(UTC).isoformat()

    def to_dict(self) -> dict[str, Any]:
        return {
            "task_id": self.task_id,
            "event_type": self.event_type,
            "sequence": self.sequence,
            "timestamp": self.timestamp,
            "payload": dict(self.payload),
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> TaskEvent:
        return cls(
            task_id=data["task_id"],
            event_type=data["event_type"],
            sequence=data["sequence"],
            timestamp=data.get("timestamp", ""),
            payload=dict(data.get("payload", {})),
        )
