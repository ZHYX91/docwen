"""Contract tests for task events and app events."""

from __future__ import annotations

import pytest

from docwen_core.events.app_events import (
    APP_STARTED,
    BATCH_COMPLETED,
    AppEvent,
)
from docwen_core.events.task_events import (
    ARTIFACT_READY,
    TASK_COMPLETED,
    TASK_FAILED,
    TASK_PROGRESS,
    TASK_STARTED,
    make_artifact_ready,
    make_diagnostic,
    make_task_cancelled,
    make_task_completed,
    make_task_failed,
    make_task_progress,
    make_task_started,
)
from docwen_core.models.task import TaskEvent

pytestmark = pytest.mark.contract


class TestTaskEventFactories:
    def test_make_task_started(self) -> None:
        evt = make_task_started("task-1", 0, file="test.md")
        assert evt.task_id == "task-1"
        assert evt.event_type == TASK_STARTED
        assert evt.sequence == 0
        assert evt.payload["file"] == "test.md"

    def test_make_task_progress(self) -> None:
        evt = make_task_progress("task-1", 5, percent=75.0, message="Converting...")
        assert evt.event_type == TASK_PROGRESS
        assert evt.payload["percent"] == 75.0
        assert evt.payload["message"] == "Converting..."

    def test_make_task_completed(self) -> None:
        evt = make_task_completed("task-1", 10, output="out.docx")
        assert evt.event_type == TASK_COMPLETED
        assert evt.payload["output"] == "out.docx"

    def test_make_task_failed(self) -> None:
        evt = make_task_failed("task-1", 3, "conversion_failed", "Broken")
        assert evt.event_type == TASK_FAILED
        assert evt.payload["error_type"] == "conversion_failed"
        assert evt.payload["message"] == "Broken"

    def test_make_task_cancelled(self) -> None:
        evt = make_task_cancelled("task-1", 7)
        assert evt.event_type == "task_cancelled"

    def test_make_artifact_ready(self) -> None:
        evt = make_artifact_ready("task-1", 8, "art-1", "output.docx")
        assert evt.event_type == ARTIFACT_READY
        assert evt.payload["artifact_id"] == "art-1"

    def test_make_diagnostic(self) -> None:
        evt = make_diagnostic("task-1", 2, "warning", "suspicious", code="W001", location="para 3")
        assert evt.event_type == "diagnostic"
        assert evt.payload["level"] == "warning"
        assert evt.payload["code"] == "W001"


class TestAppEvent:
    def test_create_and_serialize(self) -> None:
        evt = AppEvent(
            event_type=APP_STARTED,
            payload={"version": "0.1.0"},
        )
        assert evt.event_type == APP_STARTED
        assert evt.timestamp  # auto-filled
        data = evt.to_dict()
        evt2 = AppEvent.from_dict(data)
        assert evt2.event_type == APP_STARTED
        assert evt2.payload == {"version": "0.1.0"}

    def test_defaults(self) -> None:
        evt = AppEvent(event_type=BATCH_COMPLETED)
        assert evt.payload == {}
        assert evt.timestamp


class TestTaskEventOrdering:
    def test_sequence_increasing(self) -> None:
        """Events for a task should have increasing sequence numbers."""
        events: list[TaskEvent] = [
            make_task_started("task-1", 0),
            make_task_progress("task-1", 1, 50.0),
            make_task_completed("task-1", 2),
        ]
        for i in range(1, len(events)):
            assert events[i].sequence > events[i - 1].sequence

    def test_all_events_share_task_id(self) -> None:
        events: list[TaskEvent] = [
            make_task_started("abc", 0),
            make_task_progress("abc", 1, 50.0),
            make_task_failed("abc", 2, "err", "msg"),
        ]
        for evt in events:
            assert evt.task_id == "abc"
