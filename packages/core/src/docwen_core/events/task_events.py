"""Standard task event types and helper factories for ``TaskEvent``."""

from __future__ import annotations

from docwen_core.models.task import TaskEvent

# ── Event type constants ─────────────────────────────────────────────

TASK_STARTED = "task_started"
TASK_PROGRESS = "task_progress"
TASK_COMPLETED = "task_completed"
TASK_FAILED = "task_failed"
TASK_CANCELLED = "task_cancelled"
ARTIFACT_READY = "artifact_ready"
DIAGNOSTIC = "diagnostic"


# ── Factory helpers ──────────────────────────────────────────────────


def make_task_started(task_id: str, sequence: int, **payload: object) -> TaskEvent:
    """Create a ``task_started`` event."""
    return TaskEvent(
        task_id=task_id,
        event_type=TASK_STARTED,
        sequence=sequence,
        payload=dict(payload),
    )


def make_task_progress(task_id: str, sequence: int, percent: float, message: str = "") -> TaskEvent:
    """Create a ``task_progress`` event."""
    return TaskEvent(
        task_id=task_id,
        event_type=TASK_PROGRESS,
        sequence=sequence,
        payload={"percent": percent, "message": message},
    )


def make_task_completed(task_id: str, sequence: int, **payload: object) -> TaskEvent:
    """Create a ``task_completed`` event."""
    return TaskEvent(
        task_id=task_id,
        event_type=TASK_COMPLETED,
        sequence=sequence,
        payload=dict(payload),
    )


def make_task_failed(task_id: str, sequence: int, error_type: str, message: str) -> TaskEvent:
    """Create a ``task_failed`` event."""
    return TaskEvent(
        task_id=task_id,
        event_type=TASK_FAILED,
        sequence=sequence,
        payload={"error_type": error_type, "message": message},
    )


def make_task_cancelled(task_id: str, sequence: int) -> TaskEvent:
    """Create a ``task_cancelled`` event."""
    return TaskEvent(
        task_id=task_id,
        event_type=TASK_CANCELLED,
        sequence=sequence,
        payload={},
    )


def make_artifact_ready(task_id: str, sequence: int, artifact_id: str, suggested_name: str) -> TaskEvent:
    """Create an ``artifact_ready`` event."""
    return TaskEvent(
        task_id=task_id,
        event_type=ARTIFACT_READY,
        sequence=sequence,
        payload={"artifact_id": artifact_id, "suggested_name": suggested_name},
    )


def make_diagnostic(
    task_id: str,
    sequence: int,
    level: str,
    message: str,
    code: str = "",
    location: str = "",
) -> TaskEvent:
    """Create a ``diagnostic`` event."""
    return TaskEvent(
        task_id=task_id,
        event_type=DIAGNOSTIC,
        sequence=sequence,
        payload={"level": level, "message": message, "code": code, "location": location},
    )
