"""Bounded operation records independent of the editable input list."""

from __future__ import annotations

from collections import OrderedDict
from collections.abc import Callable
from copy import deepcopy
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

from PySide6.QtCore import QObject, Signal

from docwen_gui.diagnostics import DiagnosticSummary


@dataclass(frozen=True, slots=True)
class FileOutcome:
    path: str
    status: str
    error_message: str = ""
    output_path: str = ""
    warnings: tuple[str, ...] = ()
    skip_reason: str = ""
    updated_at: datetime = field(default_factory=datetime.now)
    output_paths: tuple[str, ...] = ()
    diagnostic: DiagnosticSummary | None = None


@dataclass(slots=True)
class TaskRecord:
    context: dict[str, Any]
    outcomes: dict[str, FileOutcome] = field(default_factory=dict)
    created_at: datetime = field(default_factory=datetime.now)

    @property
    def paths(self) -> tuple[str, ...]:
        return tuple(self.context.get("file_paths") or [self.context.get("file_path", "")])

    @property
    def failed_paths(self) -> list[str]:
        return [path for path in self.paths if path in self.outcomes and self.outcomes[path].status == "failed"]

    @property
    def failure_details(self) -> str:
        return "\n\n".join(
            f"{path}\n{self.outcomes[path].error_message}"
            for path in self.failed_paths
            if self.outcomes[path].error_message
        )


class TaskHistory(QObject):
    """Own retry intent and outcomes for the same bounded history as feedback."""

    changed = Signal()

    def __init__(
        self,
        limit: int = 100,
        *,
        retain_paths: Callable[[str, tuple[str, ...]], None] | None = None,
        release_owner: Callable[[str], None] | None = None,
    ) -> None:
        super().__init__()
        self._limit = limit
        self._records: OrderedDict[str, TaskRecord] = OrderedDict()
        self._retain_paths = retain_paths or (lambda _owner, _paths: None)
        self._release_owner = release_owner or (lambda _owner: None)

    def remember(self, context: dict[str, Any]) -> TaskRecord:
        operation_id = str(context.get("request_id", ""))
        if operation_id not in self._records:
            intent = deepcopy(context)
            if intent.get("file_path"):
                intent["file_path"] = Path(intent["file_path"]).as_posix()
            if intent.get("file_paths"):
                intent["file_paths"] = [Path(path).as_posix() for path in intent["file_paths"]]
            self._records[operation_id] = TaskRecord(intent)
            self._retain_paths(operation_id, self._records[operation_id].paths)
            while len(self._records) > self._limit:
                evicted_id, _evicted = self._records.popitem(last=False)
                self._release_owner(evicted_id)
            self.changed.emit()
        return self._records[operation_id]

    def get(self, operation_id: str) -> TaskRecord | None:
        return self._records.get(operation_id)

    def record(
        self,
        operation_id: str,
        path: str,
        status: str,
        error_message: str = "",
        output_path: str = "",
        *,
        warnings: tuple[str, ...] = (),
        skip_reason: str = "",
        output_paths: tuple[str, ...] = (),
        diagnostic: DiagnosticSummary | None = None,
    ) -> None:
        record = self.get(operation_id)
        if record is not None:
            path = Path(path).as_posix()
            record.outcomes[path] = FileOutcome(
                path,
                status,
                error_message,
                output_path,
                warnings,
                skip_reason,
                output_paths=output_paths or ((output_path,) if output_path else ()),
                diagnostic=diagnostic
                or DiagnosticSummary(
                    status=status,
                    output_count=len(output_paths) if output_paths else int(bool(output_path)),
                    warning_count=len(warnings),
                ),
            )
            self._sync_snapshot_retention(operation_id, record)
            self.changed.emit()

    def _sync_snapshot_retention(self, operation_id: str, record: TaskRecord) -> None:
        active = {"pending", "processing"}
        if any(path not in record.outcomes or record.outcomes[path].status in active for path in record.paths):
            return
        failed_paths = tuple(record.failed_paths)
        if failed_paths:
            self._retain_paths(operation_id, failed_paths)
        else:
            self._release_owner(operation_id)

    @property
    def entries(self) -> tuple[tuple[str, TaskRecord], ...]:
        return tuple(self._records.items())

    def clear(self) -> None:
        operation_ids = tuple(self._records)
        self._records.clear()
        for operation_id in operation_ids:
            self._release_owner(operation_id)
        self.changed.emit()

    def clear_finished(self) -> None:
        """Clearing visible history never discards an in-flight task's outcomes."""
        active = {"pending", "processing"}
        retained = OrderedDict(
            (key, record)
            for key, record in self._records.items()
            if any(path not in record.outcomes or record.outcomes[path].status in active for path in record.paths)
        )
        removed = set(self._records) - set(retained)
        self._records = retained
        for operation_id in removed:
            self._release_owner(operation_id)
        self.changed.emit()
