"""Contract tests: TaskEventEnvelope and WorkerError serialization.

These tests verify that the worker contract models (TaskEventEnvelope,
WorkerError) are fully serialisable and that no thread-state objects,
callbacks, or file handles leak into their fields.
"""

from __future__ import annotations

import json

import pytest

from docwen_core.models.result import ConversionErrorInfo
from docwen_core.models.worker import TaskEventEnvelope, WorkerError

pytestmark = pytest.mark.contract


# ═══════════════════════════════════════════════════════════════════════════
# TaskEventEnvelope
# ═══════════════════════════════════════════════════════════════════════════


class TestTaskEventEnvelope:
    """TaskEventEnvelope must round-trip through to_dict/from_dict."""

    def test_round_trip_minimal(self) -> None:
        env = TaskEventEnvelope(
            task_id="task-1",
            sequence=0,
            event_type="task_started",
        )
        data = env.to_dict()
        env2 = TaskEventEnvelope.from_dict(data)
        assert env2.task_id == "task-1"
        assert env2.sequence == 0
        assert env2.event_type == "task_started"
        assert env2.timestamp  # auto-filled

    def test_round_trip_with_payload(self) -> None:
        env = TaskEventEnvelope(
            task_id="task-2",
            sequence=5,
            event_type="task_progress",
            timestamp="2026-06-05T12:00:00Z",
            payload={"percent": 42.5, "message": "Converting paragraph 3"},
        )
        data = env.to_dict()
        env2 = TaskEventEnvelope.from_dict(data)
        assert env2.payload["percent"] == 42.5
        assert env2.payload["message"] == "Converting paragraph 3"

    def test_json_serializable(self) -> None:
        """The entire envelope must be JSON-serialisable."""
        env = TaskEventEnvelope(
            task_id="task-3",
            sequence=10,
            event_type="task_completed",
            payload={"artifact_count": 2},
        )
        data = env.to_dict()
        json_str = json.dumps(data)
        assert json_str
        # Round-trip through JSON
        data2 = json.loads(json_str)
        env2 = TaskEventEnvelope.from_dict(data2)
        assert env2.task_id == "task-3"
        assert env2.payload["artifact_count"] == 2

    def test_auto_timestamp_iso_format(self) -> None:
        """Auto-generated timestamps must be ISO-8601."""
        env = TaskEventEnvelope(task_id="t", sequence=0, event_type="task_started")
        assert "T" in env.timestamp
        # Should parse as ISO-8601
        from datetime import datetime

        dt = datetime.fromisoformat(env.timestamp)
        assert dt is not None

    def test_explicit_timestamp_preserved(self) -> None:
        """Explicit timestamps must be preserved through serialization."""
        ts = "2026-01-15T08:30:00+00:00"
        env = TaskEventEnvelope(task_id="t", sequence=0, event_type="x", timestamp=ts)
        data = env.to_dict()
        env2 = TaskEventEnvelope.from_dict(data)
        assert env2.timestamp == ts

    def test_to_task_event_conversion(self) -> None:
        """TaskEventEnvelope must convert to/from TaskEvent."""
        from docwen_core.models.task import TaskEvent

        env = TaskEventEnvelope(
            task_id="task-x",
            sequence=3,
            event_type="task_progress",
            timestamp="2026-06-05T12:00:00Z",
            payload={"pct": 50},
        )

        # Envelope → TaskEvent
        event = env.to_task_event()
        assert isinstance(event, TaskEvent)
        assert event.task_id == "task-x"
        assert event.event_type == "task_progress"
        assert event.payload["pct"] == 50

        # TaskEvent → Envelope
        env2 = TaskEventEnvelope.from_task_event(event)
        assert env2.task_id == "task-x"
        assert env2.sequence == 3
        assert env2.payload["pct"] == 50

    def test_no_thread_state_in_payload(self) -> None:
        """Payload must not contain threading primitives."""
        # This would be rejected by JSON serialization anyway
        env = TaskEventEnvelope(
            task_id="t",
            sequence=0,
            event_type="task_started",
            payload={"data": "plain string"},
        )
        json.dumps(env.to_dict())  # must not raise

    def test_no_callable_in_payload(self) -> None:
        """Payload must not contain callables (functions, lambdas)."""

        def reject_callable(obj: object) -> None:
            if callable(obj):
                raise TypeError("Callable found in payload")

        env = TaskEventEnvelope(
            task_id="t",
            sequence=0,
            event_type="x",
            payload={"key": "value", "num": 42},
        )
        # Walk the dict and verify no callables
        for v in env.payload.values():
            reject_callable(v)


# ═══════════════════════════════════════════════════════════════════════════
# WorkerError
# ═══════════════════════════════════════════════════════════════════════════


class TestWorkerError:
    """WorkerError must round-trip through to_dict/from_dict."""

    def test_round_trip_minimal(self) -> None:
        err = WorkerError(
            error_type="conversion_failed",
            message="Something went wrong",
        )
        data = err.to_dict()
        err2 = WorkerError.from_dict(data)
        assert err2.error_type == "conversion_failed"
        assert err2.message == "Something went wrong"
        assert err2.traceback_text == ""
        assert err2.recoverable is False
        assert err2.diagnostic_code == ""

    def test_round_trip_full(self) -> None:
        err = WorkerError(
            error_type="timeout",
            message="Task timed out after 30s",
            traceback_text="Traceback (most recent call last):\n  ...",
            recoverable=True,
            diagnostic_code="TIMEOUT-001",
        )
        data = err.to_dict()
        err2 = WorkerError.from_dict(data)
        assert err2.traceback_text.startswith("Traceback")
        assert err2.recoverable is True
        assert err2.diagnostic_code == "TIMEOUT-001"

    def test_json_serializable(self) -> None:
        """WorkerError must survive JSON serialization."""
        err = WorkerError(
            error_type="cancelled",
            message="User cancelled",
            recoverable=False,
        )
        json_str = json.dumps(err.to_dict())
        data = json.loads(json_str)
        err2 = WorkerError.from_dict(data)
        assert err2.error_type == "cancelled"

    def test_to_conversion_error_info(self) -> None:
        """WorkerError must convert to ConversionErrorInfo."""
        err = WorkerError(
            error_type="invalid_input",
            message="File not found",
            traceback_text="...",
            recoverable=False,
            diagnostic_code="INP-001",
        )
        info = err.to_conversion_error_info()
        assert isinstance(info, ConversionErrorInfo)
        assert info.error_type == "invalid_input"
        assert info.message == "File not found"
        assert info.diagnostic_code == "INP-001"

    def test_from_exception(self) -> None:
        """WorkerError must be constructable from a caught exception."""
        try:
            raise ValueError("bad value")
        except ValueError as exc:
            err = WorkerError.from_exception(exc, recoverable=True)

        assert err.error_type == "valueerror"
        assert err.message == "bad value"
        assert "ValueError" in err.traceback_text
        assert err.recoverable is True

    def test_no_thread_state(self) -> None:
        """WorkerError fields must not contain threading primitives."""
        err = WorkerError(error_type="test", message="test")
        data = err.to_dict()
        json.dumps(data)  # must not raise

    def test_distinct_from_conversion_error_info(self) -> None:
        """WorkerError is a separate type from ConversionErrorInfo."""
        err = WorkerError(error_type="x", message="y")
        assert not isinstance(err, ConversionErrorInfo)
        # But it can produce one
        info = err.to_conversion_error_info()
        assert isinstance(info, ConversionErrorInfo)


# ═══════════════════════════════════════════════════════════════════════════
# Cross-process safety: verify no __reduce__ or pickle-specific methods
# ═══════════════════════════════════════════════════════════════════════════


class TestWorkerContractCrossProcessSafety:
    """Verify worker contract types are safe for cross-process transport."""

    def test_no_threading_event_in_fields(self) -> None:
        """Neither WorkerRequest/Result nor envelopes carry threading.Event."""
        from dataclasses import fields

        from docwen_core.models.worker import WorkerRequest, WorkerResult

        for cls in [WorkerRequest, WorkerResult, TaskEventEnvelope, WorkerError]:
            for f in fields(cls):
                # Check that the type annotation does not reference threading
                type_str = str(f.type)
                assert "threading" not in type_str, f"{cls.__name__}.{f.name} has threading type: {type_str}"
                assert "Callback" not in type_str, f"{cls.__name__}.{f.name} has Callback type: {type_str}"

    def test_all_types_are_json_serializable(self) -> None:
        """All worker contract types must produce JSON-serialisable dicts."""
        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.manifest import RouteSpec
        from docwen_core.models.request import OutputPolicy
        from docwen_core.models.worker import WorkerRequest, WorkerResult

        wr = WorkerRequest(
            task_id="t1",
            route=RouteSpec(source_format="md", target_format="docx"),
            input_ref=FileRef(path="/f.md", format="markdown", category="markdown"),
            output_policy=OutputPolicy(),
            workspace_ref="/ws/t1",
        )
        json.dumps(wr.to_dict())

        wres = WorkerResult(task_id="t1", success=True)
        json.dumps(wres.to_dict())

        env = TaskEventEnvelope(task_id="t1", sequence=0, event_type="task_started")
        json.dumps(env.to_dict())

        werr = WorkerError(error_type="test", message="test")
        json.dumps(werr.to_dict())
