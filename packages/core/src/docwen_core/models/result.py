"""ConversionResult — the output contract for a single conversion task."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from docwen_core.models.artifact import ArtifactManifest


@dataclass(frozen=True, slots=True)
class DiagnosticRange:
    """Zero-based Unicode-code-point range with an exclusive end."""

    start: int
    end: int

    def to_dict(self) -> dict[str, int]:
        return {"start": self.start, "end": self.end}

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> DiagnosticRange:
        return cls(start=int(data["start"]), end=int(data["end"]))


@dataclass(frozen=True, slots=True)
class DiagnosticSource:
    """Immutable source identity and coordinate contract for one diagnostic."""

    input_id: str
    sha256: str
    encoding: str = "utf-8"
    coordinate_system: str = "unicode_code_point"
    offset_base: int = 0
    range_end: str = "exclusive"

    def to_dict(self) -> dict[str, Any]:
        return {
            "input_id": self.input_id,
            "sha256": self.sha256,
            "encoding": self.encoding,
            "coordinate_system": self.coordinate_system,
            "offset_base": self.offset_base,
            "range_end": self.range_end,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> DiagnosticSource:
        return cls(
            input_id=str(data["input_id"]),
            sha256=str(data["sha256"]),
            encoding=str(data["encoding"]),
            coordinate_system=str(data["coordinate_system"]),
            offset_base=int(data["offset_base"]),
            range_end=str(data["range_end"]),
        )


@dataclass(frozen=True, slots=True)
class DiagnosticTextEdit:
    """One revision-bound source edit in a diagnostic fix."""

    range: DiagnosticRange
    replacement: str

    def to_dict(self) -> dict[str, Any]:
        return {"range": self.range.to_dict(), "replacement": self.replacement}

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> DiagnosticTextEdit:
        return cls(
            range=DiagnosticRange.from_dict(data["range"]),
            replacement=str(data["replacement"]),
        )


@dataclass(frozen=True, slots=True)
class DiagnosticFix:
    """Stable fix identity plus ordered, non-overlapping source edits."""

    fix_id: str
    edits: tuple[DiagnosticTextEdit, ...]

    def to_dict(self) -> dict[str, Any]:
        return {"fix_id": self.fix_id, "edits": [edit.to_dict() for edit in self.edits]}

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> DiagnosticFix:
        return cls(
            fix_id=str(data["fix_id"]),
            edits=tuple(DiagnosticTextEdit.from_dict(edit) for edit in data["edits"]),
        )


@dataclass(slots=True)
class ConversionDiagnostic:
    """A single diagnostic message produced during conversion."""

    level: str
    """Severity: ``"info"``, ``"warning"``, ``"error"``."""

    message: str
    """Human-readable message."""

    code: str = ""
    """Machine-readable diagnostic code."""

    location: str = ""
    """Optional source location (e.g. ``"paragraph 3"``, ``"sheet:Sheet1"``)."""

    artifact_id: str | None = None
    """Optional Bundle artifact to which this diagnostic is bound."""

    evidence_schema: str | None = None
    """Optional complete Machine diagnostic evidence schema identifier."""

    source: DiagnosticSource | None = None
    """Accepted source input identity for a source-applicable diagnostic."""

    range: DiagnosticRange | None = None
    """Primary source range under ``source``'s coordinate contract."""

    related_ranges: tuple[DiagnosticRange, ...] = ()
    """Additional non-primary source ranges in deterministic order."""

    fixes: tuple[DiagnosticFix, ...] = ()
    """Bounded, revision-safe fixes for the exact accepted source."""

    def to_dict(self) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "level": self.level,
            "message": self.message,
            "code": self.code,
            "location": self.location,
        }
        if self.artifact_id is not None:
            payload["artifact_id"] = self.artifact_id
        if self.evidence_schema is not None:
            if self.source is None or self.range is None:
                raise ValueError("diagnostic evidence requires source and range")
            payload.update(
                {
                    "evidence_schema": self.evidence_schema,
                    "source": self.source.to_dict(),
                    "range": self.range.to_dict(),
                    "related_ranges": [item.to_dict() for item in self.related_ranges],
                    "fixes": [item.to_dict() for item in self.fixes],
                }
            )
        return payload

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> ConversionDiagnostic:
        return cls(
            level=data["level"],
            message=data["message"],
            code=data.get("code", ""),
            location=data.get("location", ""),
            artifact_id=data.get("artifact_id"),
            evidence_schema=data.get("evidence_schema"),
            source=DiagnosticSource.from_dict(data["source"]) if data.get("source") is not None else None,
            range=DiagnosticRange.from_dict(data["range"]) if data.get("range") is not None else None,
            related_ranges=tuple(DiagnosticRange.from_dict(item) for item in data.get("related_ranges", [])),
            fixes=tuple(DiagnosticFix.from_dict(item) for item in data.get("fixes", [])),
        )


@dataclass(slots=True)
class ConversionErrorInfo:
    """Structured error information from a failed conversion.

    This is a **data class** for serialisable error payloads.
    It is distinct from the exception class
    ``docwen_core.errors.ConversionError`` which is used for ``raise``.
    """

    error_type: str
    """Error category: ``"invalid_input"``, ``"conversion_failed"``, ``"timeout"``, etc."""

    message: str
    """Human-readable error message."""

    traceback_text: str = ""
    """Optional traceback (for debugging)."""

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
    def from_dict(cls, data: dict[str, Any]) -> ConversionErrorInfo:
        return cls(
            error_type=data["error_type"],
            message=data["message"],
            traceback_text=data.get("traceback_text", ""),
            recoverable=data.get("recoverable", False),
            diagnostic_code=data.get("diagnostic_code", ""),
        )


@dataclass(slots=True)
class ConversionMetrics:
    """Timing and size metrics for a conversion."""

    duration_ms: float = 0.0
    """Wall-clock duration in milliseconds.  ``float`` allows sub-ms precision
    for very fast conversions (e.g. < 1 ms text-only transforms)."""

    input_bytes: int = 0
    """Total input size in bytes."""

    output_bytes: int = 0
    """Total output size in bytes."""

    extra: dict[str, Any] = field(default_factory=dict)
    """Additional metrics (plugin-specific)."""

    def to_dict(self) -> dict[str, Any]:
        return {
            "duration_ms": self.duration_ms,
            "input_bytes": self.input_bytes,
            "output_bytes": self.output_bytes,
            "extra": dict(self.extra),
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> ConversionMetrics:
        return cls(
            duration_ms=data.get("duration_ms", 0.0),
            input_bytes=data.get("input_bytes", 0),
            output_bytes=data.get("output_bytes", 0),
            extra=dict(data.get("extra", {})),
        )


@dataclass(slots=True)
class ConversionResult:
    """The result of a single conversion task.

    Returned by the runtime to the application layer after finalisation.
    """

    task_id: str
    """The task id this result belongs to."""

    success: bool
    """``True`` if conversion completed without fatal error."""

    artifacts: list[ArtifactManifest] = field(default_factory=list)
    """All artifacts produced (staging paths)."""

    diagnostics: list[ConversionDiagnostic] = field(default_factory=list)
    """Diagnostics collected during conversion."""

    error: ConversionErrorInfo | None = None
    """Structured error if ``success`` is ``False``."""

    metrics: ConversionMetrics = field(default_factory=ConversionMetrics)
    """Timing and size metrics."""

    def to_dict(self) -> dict[str, Any]:
        return {
            "task_id": self.task_id,
            "success": self.success,
            "artifacts": [a.to_dict() for a in self.artifacts],
            "diagnostics": [d.to_dict() for d in self.diagnostics],
            "error": self.error.to_dict() if self.error else None,
            "metrics": self.metrics.to_dict(),
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> ConversionResult:
        return cls(
            task_id=data["task_id"],
            success=data["success"],
            artifacts=[ArtifactManifest.from_dict(a) for a in data.get("artifacts", [])],
            diagnostics=[ConversionDiagnostic.from_dict(d) for d in data.get("diagnostics", [])],
            error=ConversionErrorInfo.from_dict(data["error"]) if data.get("error") else None,
            metrics=ConversionMetrics.from_dict(data.get("metrics", {})),
        )
