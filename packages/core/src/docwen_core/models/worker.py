"""Worker contract models — serialisable messages for multi-process workers.

All models in this module MUST be fully serialisable (``to_dict``/``from_dict``).
No file handles, thread-state objects, Qt objects, callbacks, or
``threading.Event`` may appear in any field.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import UTC, datetime
from typing import Any

from docwen_core.models.artifact import ArtifactManifest
from docwen_core.models.file_ref import FileRef
from docwen_core.models.manifest import RouteSpec
from docwen_core.models.request import OutputPolicy
from docwen_core.models.result import ConversionDiagnostic, ConversionErrorInfo, ConversionMetrics


@dataclass(slots=True)
class WorkerRequest:
    """Message sent from the main process to a worker process.

    Everything in this message must be serialisable — no file handles,
    thread-state, Qt objects, or callbacks.

    **Single-file design**: each ``WorkerRequest`` carries exactly one
    ``input_ref``.  The runtime is responsible for splitting a
    ``ConversionRequest`` with multiple ``input_refs`` into individual
    ``WorkerRequest`` instances (fan-out).  For actions that genuinely
    require multiple inputs (e.g. merge, compare), the runtime's
    route resolver must provide a combined input path or extend the
    contract — this is a runtime concern, not a core model concern.

    **Cancellation**: cancellation signals are delivered through a
    separate cross-process channel (e.g. ``multiprocessing.Event``,
    file sentinel).  The ``CancellationToken`` from core MUST NOT be
    serialised into ``WorkerRequest`` because it contains a
    ``threading.Lock`` that cannot cross process boundaries.
    """

    task_id: str
    """The task id this request belongs to."""

    route: RouteSpec
    """The specific route the worker should execute."""

    input_ref: FileRef
    """The single input file to convert (one file per worker)."""

    output_policy: OutputPolicy = field(default_factory=OutputPolicy)
    """Output placement policy."""

    typed_options: dict[str, Any] = field(default_factory=dict)
    """Plugin-specific typed options."""

    config_snapshot: dict[str, Any] = field(default_factory=dict)
    """Read-only config snapshot from the main process."""

    workspace_ref: str = ""
    """Path to the workspace directory assigned to this task."""

    environment: dict[str, str] = field(default_factory=dict)
    """Selected environment variables to forward to the worker."""

    def to_dict(self) -> dict[str, Any]:
        return {
            "task_id": self.task_id,
            "route": self.route.to_dict(),
            "input_ref": self.input_ref.to_dict(),
            "output_policy": self.output_policy.to_dict(),
            "typed_options": dict(self.typed_options),
            "config_snapshot": dict(self.config_snapshot),
            "workspace_ref": self.workspace_ref,
            "environment": dict(self.environment),
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> WorkerRequest:
        return cls(
            task_id=data["task_id"],
            route=RouteSpec.from_dict(data["route"]),
            input_ref=FileRef.from_dict(data["input_ref"]),
            output_policy=OutputPolicy.from_dict(data.get("output_policy", {})),
            typed_options=dict(data.get("typed_options", {})),
            config_snapshot=dict(data.get("config_snapshot", {})),
            workspace_ref=data.get("workspace_ref", ""),
            environment=dict(data.get("environment", {})),
        )


@dataclass(slots=True)
class WorkerResult:
    """Message sent from a worker back to the main process.

    Must be fully serialisable.
    """

    task_id: str
    """The task id this result belongs to."""

    success: bool
    """``True`` if conversion completed without fatal error."""

    artifacts: list[ArtifactManifest] = field(default_factory=list)
    """Artifacts produced in staging."""

    diagnostics: list[ConversionDiagnostic] = field(default_factory=list)
    """Diagnostics collected during conversion."""

    metrics: ConversionMetrics = field(default_factory=ConversionMetrics)
    """Timing and size metrics."""

    error: ConversionErrorInfo | None = None
    """Structured error if ``success`` is ``False``."""

    def to_dict(self) -> dict[str, Any]:
        return {
            "task_id": self.task_id,
            "success": self.success,
            "artifacts": [a.to_dict() for a in self.artifacts],
            "diagnostics": [d.to_dict() for d in self.diagnostics],
            "metrics": self.metrics.to_dict(),
            "error": self.error.to_dict() if self.error else None,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> WorkerResult:
        return cls(
            task_id=data["task_id"],
            success=data["success"],
            artifacts=[ArtifactManifest.from_dict(a) for a in data.get("artifacts", [])],
            diagnostics=[ConversionDiagnostic.from_dict(d) for d in data.get("diagnostics", [])],
            metrics=ConversionMetrics.from_dict(data.get("metrics", {})),
            error=ConversionErrorInfo.from_dict(data["error"]) if data.get("error") else None,
        )


# ═══════════════════════════════════════════════════════════════════════════
# TaskEventEnvelope — serialisable event for cross-process delivery
# ═══════════════════════════════════════════════════════════════════════════


@dataclass(slots=True)
class TaskEventEnvelope:
    """A serialisable event envelope sent from worker to main process.

    This is the cross-process counterpart to ``TaskEvent``.  While
    ``TaskEvent`` is the core domain event, ``TaskEventEnvelope`` is
    explicitly designed for the worker communication channel:
    it carries the same event data but is guaranteed to contain only
    plain JSON-serialisable types (no ``datetime`` objects,
    ``threading`` primitives, or callbacks).

    The runtime's IPC layer serialises these envelopes for delivery
    across the process boundary, and the application layer deserialises
    them back into ``TaskEvent`` instances for domain use.
    """

    task_id: str
    """The task that emitted this event."""

    sequence: int
    """Monotonically increasing sequence number for ordering."""

    event_type: str
    """Event type discriminator (``"task_started"``, ``"task_progress"``, etc.)."""

    timestamp: str = ""
    """ISO-8601 UTC timestamp.  Auto-filled if empty."""

    payload: dict[str, Any] = field(default_factory=dict)
    """Type-specific payload (must be JSON-serialisable)."""

    def __post_init__(self) -> None:
        if not self.timestamp:
            self.timestamp = datetime.now(UTC).isoformat()

    def to_dict(self) -> dict[str, Any]:
        return {
            "task_id": self.task_id,
            "sequence": self.sequence,
            "event_type": self.event_type,
            "timestamp": self.timestamp,
            "payload": dict(self.payload),
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> TaskEventEnvelope:
        return cls(
            task_id=data["task_id"],
            sequence=data["sequence"],
            event_type=data["event_type"],
            timestamp=data.get("timestamp", ""),
            payload=dict(data.get("payload", {})),
        )

    def to_task_event(self) -> Any:
        """Convert to a ``TaskEvent`` for domain-layer consumption.

        Late import avoids a circular dependency (``TaskEvent`` does not
        import ``TaskEventEnvelope``).
        """
        from docwen_core.models.task import TaskEvent

        return TaskEvent(
            task_id=self.task_id,
            event_type=self.event_type,
            sequence=self.sequence,
            timestamp=self.timestamp,
            payload=dict(self.payload),
        )

    @classmethod
    def from_task_event(cls, event: Any) -> TaskEventEnvelope:
        """Create an envelope from a ``TaskEvent``."""
        return cls(
            task_id=event.task_id,
            sequence=event.sequence,
            event_type=event.event_type,
            timestamp=event.timestamp,
            payload=dict(event.payload),
        )


# ═══════════════════════════════════════════════════════════════════════════
# WorkerError — serialisable error payload for cross-process delivery
# ═══════════════════════════════════════════════════════════════════════════


@dataclass(slots=True)
class WorkerError:
    """A serialisable error container for cross-process error reporting.

    Distinct from ``ConversionErrorInfo`` (which is the domain-layer error
    payload stored inside ``ConversionResult``).  ``WorkerError`` is the
    **worker-contract** error type: it is designed to survive pickling /
    JSON serialisation across a process boundary and carries only plain
    data — no exception objects, traceback references, or thread-state.

    This type is part of the worker-contract payloads used for cross-process
    delivery.
    """

    error_type: str
    """Error category: ``"invalid_input"``, ``"conversion_failed"``,
    ``"timeout"``, ``"cancelled"``, ``"plugin_error"``, etc."""

    message: str
    """Human-readable error message."""

    traceback_text: str = ""
    """Stringified traceback (not a traceback object)."""

    recoverable: bool = False
    """Whether the task could be retried."""

    diagnostic_code: str = ""
    """Machine-readable error code."""

    def to_dict(self) -> dict[str, Any]:
        return {
            "error_type": self.error_type,
            "message": self.message,
            "traceback_text": self.traceback_text,
            "recoverable": self.recoverable,
            "diagnostic_code": self.diagnostic_code,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> WorkerError:
        return cls(
            error_type=data["error_type"],
            message=data["message"],
            traceback_text=data.get("traceback_text", ""),
            recoverable=data.get("recoverable", False),
            diagnostic_code=data.get("diagnostic_code", ""),
        )

    def to_conversion_error_info(self) -> ConversionErrorInfo:
        """Convert to a ``ConversionErrorInfo`` for domain-layer use."""
        return ConversionErrorInfo(
            error_type=self.error_type,
            message=self.message,
            traceback_text=self.traceback_text,
            recoverable=self.recoverable,
            diagnostic_code=self.diagnostic_code,
        )

    @classmethod
    def from_exception(cls, exc: Exception, recoverable: bool = False) -> WorkerError:
        """Create a ``WorkerError`` from a caught exception.

        Args:
            exc: The caught exception.
            recoverable: Whether the task can be retried.
        """
        import traceback

        return cls(
            error_type=type(exc).__name__.lower(),
            message=str(exc),
            traceback_text=traceback.format_exc(),
            recoverable=recoverable,
            diagnostic_code="",
        )
