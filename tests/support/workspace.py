"""Fake implementation of WorkspaceHandle protocol."""

from __future__ import annotations

from typing import Any


class FakeWorkspaceHandle:
    """Fake workspace — in-memory staging directory simulation.

    Does not touch the real filesystem. Artifacts are stored in memory.
    """

    def __init__(self, input_path: str, staging_dir: str, input_refs: tuple[Any, ...] = ()) -> None:
        self._input_path = input_path
        self._staging_dir = staging_dir
        self._input_refs = input_refs
        self._artifacts: list[Any] = []
        self._next_id = 0

    @property
    def input_path(self) -> str:
        return self._input_path

    @property
    def staging_dir(self) -> str:
        return self._staging_dir

    def input_resources(self, role: str | None = None) -> tuple[Any, ...]:
        if role is None:
            return self._input_refs
        return tuple(item for item in self._input_refs if getattr(item, "input_role", None) == role)

    def resource_by_logical_path(self, logical_path: str) -> Any | None:
        return next(
            (item for item in self._input_refs if getattr(item, "logical_path", None) == logical_path),
            None,
        )

    def create_artifact_path(self, kind: str, suffix: str) -> str:
        self._next_id += 1
        return f"{self._staging_dir}/artifact_{self._next_id}{suffix}"

    def add_artifact(self, manifest: Any) -> None:
        self._artifacts.append(manifest)

    @property
    def registered_artifacts(self) -> list[Any]:
        return list(self._artifacts)
