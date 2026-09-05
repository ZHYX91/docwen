"""Bounded operation records independent of the editable input list."""

from __future__ import annotations

from collections import OrderedDict
from copy import deepcopy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass(frozen=True, slots=True)
class FileOutcome:
    path: str
    status: str
    error_message: str = ""


@dataclass(slots=True)
class TaskRecord:
    context: dict[str, Any]
    outcomes: dict[str, FileOutcome] = field(default_factory=dict)

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


class TaskHistory:
    """Own retry intent and outcomes for the same bounded history as feedback."""

    def __init__(self, limit: int = 100) -> None:
        self._limit = limit
        self._records: OrderedDict[str, TaskRecord] = OrderedDict()

    def remember(self, context: dict[str, Any]) -> TaskRecord:
        operation_id = str(context.get("request_id", ""))
        if operation_id not in self._records:
            intent = deepcopy(context)
            if intent.get("file_path"):
                intent["file_path"] = Path(intent["file_path"]).as_posix()
            if intent.get("file_paths"):
                intent["file_paths"] = [Path(path).as_posix() for path in intent["file_paths"]]
            self._records[operation_id] = TaskRecord(intent)
            while len(self._records) > self._limit:
                self._records.popitem(last=False)
        return self._records[operation_id]

    def get(self, operation_id: str) -> TaskRecord | None:
        return self._records.get(operation_id)

    def record(self, operation_id: str, path: str, status: str, error_message: str = "") -> None:
        record = self.get(operation_id)
        if record is not None:
            path = Path(path).as_posix()
            record.outcomes[path] = FileOutcome(path, status, error_message)

    def clear(self) -> None:
        self._records.clear()
