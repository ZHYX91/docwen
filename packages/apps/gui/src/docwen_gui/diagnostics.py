"""Bounded diagnostic facts safe to preview and copy without document data."""

from __future__ import annotations

import json
from dataclasses import dataclass
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from docwen_core.models.result import ConversionResult

_STATUSES = frozenset(
    {
        "info",
        "success",
        "warning",
        "error",
        "pending",
        "processing",
        "completed",
        "partial",
        "failed",
        "skipped",
        "cancelled",
    }
)
_ERROR_CATEGORIES = frozenset(
    {
        "invalid_input",
        "unsupported_route",
        "conversion_failed",
        "output_finalization_failed",
        "dependency_missing",
        "dependency_unavailable",
        "timeout",
        "cancelled",
        "skipped",
        "permission_denied",
        "file_not_found",
        "configuration_error",
        "internal_error",
    }
)
_EXCEPTION_TYPES = frozenset(
    {
        "OSError",
        "PermissionError",
        "FileNotFoundError",
        "FileExistsError",
        "TimeoutError",
        "ValueError",
        "TypeError",
        "RuntimeError",
        "UnicodeError",
        "UnicodeDecodeError",
    }
)


def _count(value: int) -> int:
    return min(value, 1_000_000) if type(value) is int and value >= 0 else 0


@dataclass(frozen=True, slots=True)
class DiagnosticSummary:
    """Only reviewed finite labels, counts and a reported retry flag leave the UI.

    No message, exception repr, path, source hash, operation ID, option or arbitrary
    metadata is accepted here. Unknown categories remain unknown; syntax alone
    does not make an arbitrary error code safe to copy.
    """

    status: str = "unknown"
    error_category: str = ""
    exception_type: str = ""
    output_count: int = 0
    warning_count: int = 0
    reported_recoverable: bool | None = None
    succeeded_count: int | None = None
    failed_count: int | None = None
    cancelled_count: int | None = None

    def to_text(self) -> str:
        payload: dict[str, str | int | bool] = {
            "schema": "docwen.gui-diagnostic.v1",
            "status": self.status if self.status in _STATUSES else "unknown",
            "output_count": _count(self.output_count),
            "warning_count": _count(self.warning_count),
        }
        if self.error_category:
            payload["error_category"] = self.error_category if self.error_category in _ERROR_CATEGORIES else "unknown"
        if self.exception_type:
            payload["exception_type"] = self.exception_type if self.exception_type in _EXCEPTION_TYPES else "unknown"
        if type(self.reported_recoverable) is bool:
            payload["reported_recoverable"] = self.reported_recoverable
        for key, value in (
            ("succeeded_count", self.succeeded_count),
            ("failed_count", self.failed_count),
            ("cancelled_count", self.cancelled_count),
        ):
            if value is not None:
                payload[key] = _count(value)
        return json.dumps(payload, ensure_ascii=False, indent=2)

    @classmethod
    def from_result(cls, result: ConversionResult, *, output_count: int) -> DiagnosticSummary:
        error = result.error
        category = error.error_type if error is not None else ""
        return cls(
            status="completed" if result.success else category if category in {"cancelled", "skipped"} else "failed",
            error_category=category,
            output_count=output_count,
            warning_count=sum(item.level == "warning" for item in result.diagnostics),
            reported_recoverable=error.recoverable if error is not None else None,
        )

    @classmethod
    def from_exception(cls, exc: BaseException) -> DiagnosticSummary:
        return cls(status="error", exception_type=type(exc).__name__)
