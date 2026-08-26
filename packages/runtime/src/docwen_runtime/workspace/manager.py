"""WorkspaceManager — creates and manages per-task workspace directories.

Each task gets its own workspace with:
- input/   — symlink or copy of input file (convenience)
- staging/ — where plugins write intermediate/output artifacts
- manifest.json — artifact manifest written by the runtime after conversion
"""

from __future__ import annotations

import hashlib
import os
import shutil
import sys
import tempfile
import threading
from dataclasses import replace
from pathlib import Path
from typing import TYPE_CHECKING

from docwen_core.models.file_ref import FileRef
from docwen_runtime.path_io import filesystem_path

if TYPE_CHECKING:
    from docwen_core.models.artifact import ArtifactManifest


class WorkspaceHandle:
    """Concrete implementation of the ``WorkspaceHandle`` protocol.

    Provides plugins with staging paths and artifact registration.
    Plugins interact with this handle through the
    ``docwen_core.protocols.execution_context.WorkspaceHandle`` Protocol.
    """

    def __init__(self, input_path: str, staging_dir: str, input_refs: tuple[FileRef, ...] = ()) -> None:
        self._input_path = input_path
        self._staging_dir = staging_dir
        self._input_refs = input_refs
        self._input_by_logical_path = {item.logical_path: item for item in input_refs if item.logical_path}
        self._artifacts: list[ArtifactManifest] = []
        self._counter = 0

    @property
    def input_path(self) -> str:
        return self._input_path

    @property
    def staging_dir(self) -> str:
        return self._staging_dir

    def input_resources(self, role: str | None = None) -> tuple[FileRef, ...]:
        if role is None:
            return self._input_refs
        return tuple(item for item in self._input_refs if item.input_role == role)

    def resource_by_logical_path(self, logical_path: str) -> FileRef | None:
        return self._input_by_logical_path.get(logical_path)

    def create_artifact_path(self, kind: str, suffix: str) -> str:
        """Create and return a writable path inside the staging directory.

        The path is guaranteed to be unique within this workspace.
        The caller (plugin) is responsible for actually writing the file.
        """
        self._counter += 1
        safe_kind = kind.replace("/", "_").replace("\\", "_")
        filename = f"{safe_kind}_{self._counter}{suffix}"
        return os.path.join(self._staging_dir, filename)

    def add_artifact(self, manifest: ArtifactManifest) -> None:
        """Register an artifact manifest."""
        self._artifacts.append(manifest)

    @property
    def registered_artifacts(self) -> list[ArtifactManifest]:
        """Return all registered artifacts (copy)."""
        return list(self._artifacts)


class WorkspaceManager:
    """Creates and cleans up per-task workspace directories.

    Workspaces are created under a configurable root directory.
    Each workspace is a unique sub-directory identified by task id.

    Lifecycle:
    1. ``create(task_id, input_path)`` → ``WorkspaceHandle``
    2. Plugin writes to staging via handle
    3. Output finalizer reads artifacts and places final output
    4. ``cleanup(task_id)`` removes the workspace directory
    """

    def __init__(self, root_dir: str | os.PathLike[str]) -> None:
        """Initialize the workspace manager.

        Args:
            root_dir: Absolute, non-filesystem-root directory for workspaces.

        Raises:
            ValueError: If the root is relative or resolves lexically to a
                filesystem root. Callers must select an owned directory;
                silently falling back to the process temp directory is not
                allowed.
        """
        selected = Path(root_dir)
        if not selected.is_absolute():
            raise ValueError("workspace root must be absolute")
        selected = Path(os.path.normpath(os.fspath(selected)))
        if selected == Path(selected.anchor):
            raise ValueError("workspace root must not be a filesystem root")
        self._root = os.fspath(selected)
        self._workspaces: dict[str, WorkspaceHandle] = {}
        self._lock = threading.RLock()

    @property
    def root_dir(self) -> str:
        return self._root

    def create(
        self,
        task_id: str,
        input_path: str,
        input_refs: tuple[FileRef, ...] = (),
    ) -> WorkspaceHandle:
        """Create a workspace for a task.

        Args:
            task_id: Unique task identifier.
            input_path: Path to the input file.

        Returns:
            A ``WorkspaceHandle`` the plugin can use for staging writes.
        """
        with self._lock:
            if task_id in self._workspaces:
                raise ValueError(f"Workspace already exists for task: {task_id!r}")

            # Keep the caller-controlled task id out of filesystem paths.
            # Besides preventing path traversal, a short opaque name leaves
            # headroom for external Office applications with small path budgets.
            os.makedirs(self._root, exist_ok=True)
            ws_dir = tempfile.mkdtemp(prefix="dw-", dir=self._root)
            staging_dir = os.path.join(ws_dir, "staging")
            try:
                os.makedirs(staging_dir)
                declared_refs, declared_input_path = self._materialize_declared_inputs(ws_dir, input_refs)
            except Exception:
                shutil.rmtree(ws_dir, ignore_errors=True)
                raise

            handle = WorkspaceHandle(declared_input_path or input_path, staging_dir, declared_refs or input_refs)
            self._workspaces[task_id] = handle
            return handle

    @staticmethod
    def _materialize_declared_inputs(
        workspace_dir: str,
        input_refs: tuple[FileRef, ...],
    ) -> tuple[tuple[FileRef, ...], str | None]:
        """Copy typed inputs into their isolated request virtual root."""
        if not input_refs or not any(item.logical_path for item in input_refs):
            return (), None
        if any(not item.logical_path for item in input_refs):
            raise ValueError("typed input sets require logical_path on every input")
        input_root = Path(workspace_dir, "inputs")
        input_root.mkdir()
        materialized: list[FileRef] = []
        source_path: str | None = None
        logical_paths: set[str] = set()
        for index, item in enumerate(input_refs):
            segments = item.logical_path.split("/")
            if (
                not item.logical_path
                or len(item.logical_path) > 1024
                or item.logical_path.startswith("/")
                or "\\" in item.logical_path
                or "\x00" in item.logical_path
                or ":" in segments[0]
                or any(segment in {"", ".", ".."} for segment in segments)
            ):
                raise ValueError("invalid typed input logical_path")
            if item.logical_path in logical_paths:
                raise ValueError("duplicate typed input logical_path")
            logical_paths.add(item.logical_path)
            source = filesystem_path(item.path, force_extended=sys.platform == "win32")
            if WorkspaceManager._path_traverses_link_or_junction(source):
                raise ValueError("typed input must not be a link or junction")
            if not source.is_file():
                raise ValueError("typed input must be an existing regular file")
            suffix = source.suffix if len(source.suffix) <= 32 else ""
            destination = input_root / f"input-{index:04d}{suffix}"
            shutil.copy2(source, destination)
            expected_sha = item.metadata.get("machine_input_sha256")
            expected_size = item.metadata.get("machine_input_size_bytes")
            if isinstance(expected_sha, str) and isinstance(expected_size, int):
                digest = hashlib.sha256()
                size_bytes = 0
                with destination.open("rb") as stream:
                    while chunk := stream.read(1024 * 1024):
                        size_bytes += len(chunk)
                        digest.update(chunk)
                if size_bytes != expected_size or digest.hexdigest() != expected_sha:
                    raise ValueError("typed input copy failed integrity verification")
            copied = replace(item, path=str(destination))
            materialized.append(copied)
            if item.input_role in {"source", "neutral_document"} and source_path is None:
                source_path = str(destination)
        if source_path is None:
            raise ValueError("typed input set has no primary document input")
        return tuple(materialized), source_path

    @staticmethod
    def _is_link_or_junction(path: Path) -> bool:
        if path.is_symlink():
            return True
        is_junction = getattr(path, "is_junction", None)
        return bool(is_junction()) if callable(is_junction) else False

    @classmethod
    def _path_traverses_link_or_junction(cls, path: Path) -> bool:
        current = path
        while True:
            if cls._is_link_or_junction(current):
                return True
            parent = current.parent
            if parent == current:
                return False
            current = parent

    def get(self, task_id: str) -> WorkspaceHandle | None:
        """Return the workspace handle for *task_id*, or ``None``."""
        with self._lock:
            return self._workspaces.get(task_id)

    def cleanup(self, task_id: str) -> None:
        """Remove the workspace directory for *task_id*.

        No-op if the task has no workspace or the directory is already gone.
        """
        with self._lock:
            handle = self._workspaces.pop(task_id, None)
        if handle is None:
            return
        ws_dir = os.path.dirname(handle.staging_dir)
        if os.path.isdir(ws_dir):
            shutil.rmtree(ws_dir, ignore_errors=True)

    def cleanup_all(self) -> None:
        """Remove all workspace directories."""
        with self._lock:
            task_ids = tuple(self._workspaces)
        for task_id in task_ids:
            self.cleanup(task_id)

    def __len__(self) -> int:
        with self._lock:
            return len(self._workspaces)
