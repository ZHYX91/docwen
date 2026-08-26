"""OutputFinalizer — reads staging artifacts and places them at the final output path.

This is the ONLY component allowed to write to the final output directory.
Plugins MUST NOT call any method on this class directly.
"""

from __future__ import annotations

import contextlib
import errno
import hashlib
import json
import os
import shutil
import sys
import tempfile
import threading
import time
import weakref
from dataclasses import dataclass, replace
from pathlib import Path
from typing import TYPE_CHECKING, Any

from docwen_core.errors import CancellationRequested
from docwen_core.models.artifact import ArtifactManifest
from docwen_core.models.result import (
    ConversionDiagnostic,
    ConversionErrorInfo,
    ConversionMetrics,
    ConversionResult,
)
from docwen_runtime.output.document_node import (
    DOCUMENT_NODE_MANIFEST_MEDIA_TYPE,
    DocumentNodeLayoutPlan,
    has_markdown_artifacts,
    plan_document_node_layout,
    relocated_markdown_bytes,
)
from docwen_runtime.path_io import filesystem_path

if TYPE_CHECKING:
    from docwen_core.models.request import OutputPolicy
    from docwen_core.protocols.execution_context import CancellationTokenView


_TEMP_PREFIX = ".__docwen-finalizer-"
_LOCK_POLL_SECONDS = 0.05
_COPY_CHUNK_BYTES = 1024 * 1024
_STALE_TEMP_MIN_AGE_SECONDS = 24 * 60 * 60


@dataclass
class _PreparedArtifact:
    """One artifact prepared for a no-torn-write commit."""

    artifact: ArtifactManifest
    suggested_name: str
    destination: str
    rename_base: str | None = None
    temp_path: str | None = None
    reuse: ArtifactManifest | None = None
    skip_existing: bool = False


if sys.platform == "win32":
    import msvcrt  # type: ignore[import-untyped]

    def _try_lock_file(file_descriptor: int) -> None:
        os.lseek(file_descriptor, 0, os.SEEK_SET)
        msvcrt.locking(file_descriptor, msvcrt.LK_NBLCK, 1)

    def _unlock_file(file_descriptor: int) -> None:
        os.lseek(file_descriptor, 0, os.SEEK_SET)
        msvcrt.locking(file_descriptor, msvcrt.LK_UNLCK, 1)

else:
    import fcntl

    def _try_lock_file(file_descriptor: int) -> None:
        fcntl.flock(file_descriptor, fcntl.LOCK_EX | fcntl.LOCK_NB)

    def _unlock_file(file_descriptor: int) -> None:
        fcntl.flock(file_descriptor, fcntl.LOCK_UN)


class OutputFinalizer:
    """Reads staging artifacts and performs final placement.

    Responsibilities:
    - Apply output policy (output_dir, date_subfolder, overwrite_mode).
    - Handle file name collisions (rename, overwrite, skip).
    - Copy/move artifacts from staging to final location.
    - Return a ``ConversionResult`` with final paths.

    This is the single choke-point for all final output writes.
    """

    _output_locks_guard = threading.Lock()
    _output_locks: weakref.WeakValueDictionary[str, Any] = weakref.WeakValueDictionary()

    def resolve_output_dir(self, policy: OutputPolicy, input_path: str = "") -> str:
        """Resolve the exact final directory for trusted runtime sidecars."""
        return os.path.abspath(self._resolve_output_dir(policy, input_path))

    def finalize(
        self,
        task_id: str,
        artifacts: list[ArtifactManifest],
        policy: OutputPolicy,
        *,
        input_path: str = "",
        duration_ms: float = 0.0,
        input_bytes: int = 0,
        cancellation: CancellationTokenView | None = None,
    ) -> ConversionResult:
        """Finalize a set of artifacts.

        Args:
            task_id: The task identifier.
            artifacts: Artifact manifests from the plugin (staging paths).
            policy: Output placement policy.
            input_path: Original input file path (for ``same_dir`` mode).
            duration_ms: Conversion duration for metrics.
            input_bytes: Input size for metrics.

        Returns:
            A ``ConversionResult`` with final (not staging) artifact paths.
        """
        self._check_cancellation(cancellation)
        output_dir = self._resolve_output_dir(policy, input_path)
        node_plan: DocumentNodeLayoutPlan | None = None
        if has_markdown_artifacts(artifacts):
            in_place_markdown = bool(
                policy.output_path
                and input_path
                and os.path.normcase(os.path.abspath(policy.output_path))
                == os.path.normcase(os.path.abspath(input_path))
            )
            if policy.output_path and not in_place_markdown:
                raise ValueError("Markdown output requires an output parent directory, not output_path")
            if in_place_markdown:
                artifacts = self._artifacts_for_policy(artifacts, policy)
            else:
                node_plan = plan_document_node_layout(
                    task_id=task_id,
                    artifacts=artifacts,
                    input_path=input_path,
                )
                artifacts = list(node_plan.artifacts)
        else:
            artifacts = self._artifacts_for_policy(artifacts, policy)
        lock_paths = self._finalization_lock_paths(output_dir, artifacts)
        with self._finalization_locks(lock_paths, cancellation):
            if node_plan is not None:
                return self._finalize_document_node_locked(
                    task_id,
                    node_plan,
                    policy,
                    output_dir=output_dir,
                    input_path=input_path,
                    duration_ms=duration_ms,
                    input_bytes=input_bytes,
                    cancellation=cancellation,
                )
            return self._finalize_locked(
                task_id,
                artifacts,
                policy,
                output_dir=output_dir,
                input_path=input_path,
                duration_ms=duration_ms,
                input_bytes=input_bytes,
                cancellation=cancellation,
            )

    def _finalize_locked(
        self,
        task_id: str,
        artifacts: list[ArtifactManifest],
        policy: OutputPolicy,
        *,
        output_dir: str,
        input_path: str,
        duration_ms: float,
        input_bytes: int,
        cancellation: CancellationTokenView | None,
    ) -> ConversionResult:
        """Finalize one artifact batch while all concrete parent locks are held."""
        self._check_cancellation(cancellation)
        filesystem_path(output_dir, force_extended=sys.platform == "win32").mkdir(
            parents=True,
            exist_ok=True,
        )

        diagnostics: list[ConversionDiagnostic] = []
        final_artifacts: list[ArtifactManifest] = []
        total_output_bytes = 0
        failed_artifacts = 0
        prepared: list[_PreparedArtifact] = []

        try:
            cleaned_parents: set[str] = set()
            for artifact in artifacts:
                try:
                    suggested = artifact.suggested_name or os.path.basename(artifact.staging_path)
                    destination, _ = self._safe_final_path(output_dir, suggested)
                    parent = os.path.dirname(destination)
                    parent_key = os.path.normcase(os.path.realpath(parent))
                    if parent_key not in cleaned_parents and self._io_path(parent).is_dir():
                        self._cleanup_stale_temps(parent)
                        cleaned_parents.add(parent_key)
                except (OSError, ValueError):
                    # Placement below owns the typed per-artifact diagnostic.
                    pass
            for artifact in artifacts:
                try:
                    self._check_cancellation(cancellation)
                    prepared.append(
                        self._prepare_artifact(
                            artifact,
                            output_dir,
                            policy.overwrite_mode,
                            input_path,
                            cancellation,
                        )
                    )
                except CancellationRequested:
                    raise
                except Exception as exc:
                    failed_artifacts += 1
                    diagnostics.append(
                        ConversionDiagnostic(
                            level="error",
                            message=(
                                f"Failed to place artifact {artifact.artifact_id!r}: {self._public_exception_text(exc)}"
                            ),
                            code="FINALIZER_PLACE_ERROR",
                        )
                    )

            # This is the batch linearization point. Cancellation wins through
            # this check; after it returns, complete artifact commits win.
            self._check_cancellation(cancellation)
            for item in prepared:
                try:
                    final_artifact, written_bytes = self._commit_prepared(
                        item,
                        output_dir,
                        policy.overwrite_mode,
                    )
                    final_artifacts.append(final_artifact)
                    total_output_bytes += written_bytes
                except Exception as exc:
                    failed_artifacts += 1
                    diagnostics.append(
                        ConversionDiagnostic(
                            level="error",
                            message=(
                                f"Failed to place artifact {item.artifact.artifact_id!r}: "
                                f"{self._public_exception_text(exc)}"
                            ),
                            code="FINALIZER_PLACE_ERROR",
                        )
                    )
        finally:
            for item in prepared:
                if item.temp_path:
                    with contextlib.suppress(OSError):
                        self._io_path(item.temp_path).unlink()

        attempted_artifacts = len(artifacts)
        placed_artifacts = len(final_artifacts)
        error: ConversionErrorInfo | None = None
        metrics_extra: dict[str, Any] = {"output_dir": output_dir}

        if attempted_artifacts == 0:
            summary_code = "FINALIZER_NO_ARTIFACTS"
            summary_message = "No output artifacts were provided for finalization"
            diagnostics.append(
                ConversionDiagnostic(
                    level="error",
                    message=summary_message,
                    code=summary_code,
                )
            )
            error = ConversionErrorInfo(
                error_type="output_finalization_failed",
                message=summary_message,
                diagnostic_code=summary_code,
            )
        elif failed_artifacts:
            summary_code = "FINALIZER_PARTIAL" if placed_artifacts else "FINALIZER_FAILED"
            summary_message = f"Placed {placed_artifacts} of {attempted_artifacts} artifact(s) in {output_dir}"
            diagnostics.append(
                ConversionDiagnostic(
                    level="error",
                    message=summary_message,
                    code=summary_code,
                )
            )
            error = ConversionErrorInfo(
                error_type="output_finalization_failed",
                message=summary_message,
                diagnostic_code=summary_code,
            )
            metrics_extra.update(
                {
                    "attempted_artifacts": attempted_artifacts,
                    "placed_artifacts": placed_artifacts,
                    "failed_artifacts": failed_artifacts,
                }
            )
        else:
            diagnostics.append(
                ConversionDiagnostic(
                    level="info",
                    message=f"Placed {placed_artifacts} artifact(s) in {output_dir}",
                    code="FINALIZER_DONE",
                )
            )

        return ConversionResult(
            task_id=task_id,
            success=error is None,
            artifacts=final_artifacts,
            diagnostics=diagnostics,
            error=error,
            metrics=ConversionMetrics(
                duration_ms=duration_ms,
                input_bytes=input_bytes,
                output_bytes=total_output_bytes,
                extra=metrics_extra,
            ),
        )

    def _finalize_document_node_locked(
        self,
        task_id: str,
        plan: DocumentNodeLayoutPlan,
        policy: OutputPolicy,
        *,
        output_dir: str,
        input_path: str,
        duration_ms: float,
        input_bytes: int,
        cancellation: CancellationTokenView | None,
    ) -> ConversionResult:
        """Publish a complete Markdown node through one directory commit."""

        self._check_cancellation(cancellation)
        output_io = self._io_path(output_dir)
        output_io.mkdir(parents=True, exist_ok=True)
        selected = plan
        collision = 0
        while self._io_path(os.path.join(output_dir, selected.root_name)).exists():
            if policy.overwrite_mode == "rename":
                collision += 1
                selected = plan.rebase_root(plan.identity.node_name(collision=collision))
                continue
            if policy.overwrite_mode in {"overwrite", "skip"}:
                break
            return self._document_node_failure(
                task_id,
                f"Document node already exists: {os.path.join(output_dir, selected.root_name)}",
                duration_ms=duration_ms,
                input_bytes=input_bytes,
                output_dir=output_dir,
                root_name=selected.root_name,
                code="DOCUMENT_NODE_EXISTS",
            )

        final_root = os.path.abspath(os.path.join(output_dir, selected.root_name))
        self._ensure_contained(output_dir, final_root)
        source_sha256 = self._sha256_if_file(input_path, cancellation)
        if policy.overwrite_mode == "skip" and self._io_path(final_root).exists():
            return self._reuse_document_node(
                task_id,
                selected,
                final_root=final_root,
                source_sha256=source_sha256,
                duration_ms=duration_ms,
                input_bytes=input_bytes,
                cancellation=cancellation,
            )

        # ``mkdtemp`` appends a private name, so reserve Win32 namespace headroom
        # even when the public output directory itself is still below MAX_PATH.
        temp_parent = filesystem_path(output_dir, force_extended=sys.platform == "win32")
        temp_root = tempfile.mkdtemp(prefix=".__docwen-node-", dir=os.fspath(temp_parent))
        backup_root: str | None = None
        committed = False
        try:
            self._ensure_contained(output_dir, temp_root)
            manifest_artifacts: list[dict[str, Any]] = []
            output_bytes = 0
            for artifact in selected.artifacts:
                self._check_cancellation(cancellation)
                if artifact.logical_path is None:
                    raise ValueError(f"Artifact {artifact.artifact_id!r} has no logical_path")
                logical = artifact.logical_path.replace("\\", "/")
                prefix = f"{selected.root_name}/"
                if not logical.startswith(prefix):
                    raise ValueError(f"Artifact escapes document node: {logical!r}")
                inner = logical[len(prefix) :]
                destination = os.path.abspath(os.path.join(temp_root, *inner.split("/")))
                self._ensure_contained(temp_root, destination)
                destination_io = self._io_path(destination)
                destination_io.parent.mkdir(parents=True, exist_ok=True)
                if artifact.media_type == "text/markdown":
                    payload = relocated_markdown_bytes(artifact, artifacts=selected.artifacts)
                    with destination_io.open("wb") as stream:
                        stream.write(payload)
                        stream.flush()
                        os.fsync(stream.fileno())
                    with contextlib.suppress(OSError):
                        shutil.copystat(self._io_path(artifact.staging_path), destination_io)
                else:
                    self._copy_to_temp(artifact.staging_path, destination, cancellation)
                size_bytes, sha256 = self._file_integrity(destination, cancellation)
                output_bytes += size_bytes
                manifest_artifacts.append(
                    {
                        "artifact_id": artifact.artifact_id,
                        "kind": artifact.kind,
                        "logical_path": logical,
                        "media_type": artifact.media_type,
                        "role": artifact.metadata.get("document_node_role", "resource"),
                        "size_bytes": size_bytes,
                        "sha256": sha256,
                    }
                )

            manifest_relative = "docwen-node.json"
            manifest_temp = os.path.join(temp_root, manifest_relative)
            manifest_document = {
                "schema": "docwen.document_node.v1",
                "task_id": task_id,
                "node_name": selected.root_name,
                "created_at": selected.identity.created_at_utc,
                "source": {
                    "name": os.path.basename(input_path) if input_path else "",
                    "format": selected.identity.source_format,
                    "sha256": source_sha256,
                },
                "artifacts": manifest_artifacts,
            }
            manifest_bytes = (
                json.dumps(manifest_document, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n"
            ).encode("utf-8")
            with self._io_path(manifest_temp).open("wb") as stream:
                stream.write(manifest_bytes)
                stream.flush()
                os.fsync(stream.fileno())
            output_bytes += len(manifest_bytes)

            self._check_cancellation(cancellation)
            if self._io_path(final_root).exists():
                if policy.overwrite_mode != "overwrite":
                    raise FileExistsError(f"Document node already exists: {final_root}")
                self._validate_owned_document_node(final_root)
                backup_root = self._unused_backup_path(final_root)
                self._ensure_contained(output_dir, backup_root)
                os.rename(os.fspath(self._io_path(final_root)), os.fspath(self._io_path(backup_root)))
            try:
                os.rename(os.fspath(self._io_path(temp_root)), os.fspath(self._io_path(final_root)))
                committed = True
                temp_root = ""
            except BaseException:
                if backup_root is not None and not self._io_path(final_root).exists():
                    os.rename(os.fspath(self._io_path(backup_root)), os.fspath(self._io_path(final_root)))
                    backup_root = None
                raise
            if backup_root is not None:
                with contextlib.suppress(OSError):
                    shutil.rmtree(self._io_path(backup_root))
                backup_root = None

            placed = [
                replace(
                    artifact,
                    staging_path=os.path.join(final_root, *artifact.logical_path.split("/")[1:]),
                    metadata={**artifact.metadata, "document_node_committed": True},
                )
                for artifact in selected.artifacts
                if artifact.logical_path is not None
            ]
            manifest_logical = f"{selected.root_name}/{manifest_relative}"
            placed.append(
                ArtifactManifest(
                    artifact_id=f"{task_id}-document-node-manifest",
                    kind="manifest",
                    staging_path=os.path.join(final_root, manifest_relative),
                    suggested_name=manifest_relative,
                    media_type=DOCUMENT_NODE_MANIFEST_MEDIA_TYPE,
                    metadata={
                        "document_node_schema": "docwen.document_node.v1",
                        "document_node_role": "manifest",
                        "node_root": selected.root_name,
                        "logical_path": manifest_logical,
                    },
                    is_primary=False,
                    logical_path=manifest_logical,
                )
            )
            return ConversionResult(
                task_id=task_id,
                success=True,
                artifacts=placed,
                diagnostics=[
                    ConversionDiagnostic(
                        level="info",
                        message=f"Published document node {final_root}",
                        code="FINALIZER_DONE",
                    )
                ],
                metrics=ConversionMetrics(
                    duration_ms=duration_ms,
                    input_bytes=input_bytes,
                    output_bytes=output_bytes,
                    extra={
                        "output_dir": output_dir,
                        "document_node_root": final_root,
                        "document_node_schema": "docwen.document_node.v1",
                    },
                ),
            )
        except CancellationRequested:
            raise
        except Exception as exc:
            return self._document_node_failure(
                task_id,
                self._public_exception_text(exc),
                duration_ms=duration_ms,
                input_bytes=input_bytes,
                output_dir=output_dir,
                root_name=selected.root_name,
            )
        finally:
            if temp_root and self._io_path(temp_root).exists():
                shutil.rmtree(self._io_path(temp_root), ignore_errors=True)
            if (
                backup_root is not None
                and self._io_path(backup_root).exists()
                and not committed
                and not self._io_path(final_root).exists()
            ):
                with contextlib.suppress(OSError):
                    os.rename(os.fspath(self._io_path(backup_root)), os.fspath(self._io_path(final_root)))

    @staticmethod
    def _document_node_failure(
        task_id: str,
        message: str,
        *,
        duration_ms: float,
        input_bytes: int,
        output_dir: str,
        root_name: str,
        code: str = "DOCUMENT_NODE_PUBLISH_FAILED",
    ) -> ConversionResult:
        return ConversionResult(
            task_id=task_id,
            success=False,
            artifacts=[],
            diagnostics=[ConversionDiagnostic(level="error", message=message, code=code)],
            error=ConversionErrorInfo(
                error_type="output_finalization_failed",
                message=message,
                diagnostic_code=code,
            ),
            metrics=ConversionMetrics(
                duration_ms=duration_ms,
                input_bytes=input_bytes,
                output_bytes=0,
                extra={"output_dir": output_dir, "document_node_root": os.path.join(output_dir, root_name)},
            ),
        )

    def _reuse_document_node(
        self,
        task_id: str,
        plan: DocumentNodeLayoutPlan,
        *,
        final_root: str,
        source_sha256: str | None,
        duration_ms: float,
        input_bytes: int,
        cancellation: CancellationTokenView | None,
    ) -> ConversionResult:
        manifest_path = self._io_path(os.path.join(final_root, "docwen-node.json"))
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            if (
                manifest.get("schema") != "docwen.document_node.v1"
                or manifest.get("source", {}).get("sha256") != source_sha256
            ):
                raise ValueError("existing node identity does not match this source")
            recorded = {
                item.get("logical_path"): item
                for item in manifest.get("artifacts", [])
                if isinstance(item, dict) and isinstance(item.get("logical_path"), str)
            }
            placed = []
            for artifact in plan.artifacts:
                if artifact.logical_path is None:
                    raise ValueError("planned artifact has no logical path")
                path = os.path.join(final_root, *artifact.logical_path.split("/")[1:])
                if not self._io_path(path).is_file():
                    raise FileNotFoundError(path)
                expected = recorded.get(artifact.logical_path)
                size_bytes, sha256 = self._file_integrity(path, cancellation)
                if expected is None or expected.get("size_bytes") != size_bytes or expected.get("sha256") != sha256:
                    raise ValueError(f"existing node artifact integrity mismatch: {artifact.logical_path}")
                placed.append(
                    replace(
                        artifact,
                        staging_path=path,
                        metadata={**artifact.metadata, "document_node_reused": True},
                    )
                )
            manifest_logical = f"{plan.root_name}/docwen-node.json"
            placed.append(
                ArtifactManifest(
                    artifact_id=f"{task_id}-document-node-manifest",
                    kind="manifest",
                    staging_path=os.fspath(manifest_path),
                    suggested_name="docwen-node.json",
                    media_type=DOCUMENT_NODE_MANIFEST_MEDIA_TYPE,
                    metadata={
                        "document_node_schema": "docwen.document_node.v1",
                        "document_node_role": "manifest",
                        "node_root": plan.root_name,
                        "logical_path": manifest_logical,
                        "document_node_reused": True,
                    },
                    is_primary=False,
                    logical_path=manifest_logical,
                )
            )
        except Exception as exc:
            return self._document_node_failure(
                task_id,
                self._public_exception_text(exc),
                duration_ms=duration_ms,
                input_bytes=input_bytes,
                output_dir=os.path.dirname(final_root),
                root_name=os.path.basename(final_root),
                code="DOCUMENT_NODE_SKIP_MISMATCH",
            )
        return ConversionResult(
            task_id=task_id,
            success=True,
            artifacts=placed,
            diagnostics=[
                ConversionDiagnostic(
                    level="info", message=f"Reused document node {final_root}", code="DOCUMENT_NODE_REUSED"
                )
            ],
            metrics=ConversionMetrics(
                duration_ms=duration_ms,
                input_bytes=input_bytes,
                output_bytes=0,
                extra={"output_dir": os.path.dirname(final_root), "document_node_root": final_root},
            ),
        )

    def _validate_owned_document_node(self, root: str) -> None:
        manifest_path = self._io_path(os.path.join(root, "docwen-node.json"))
        try:
            document = json.loads(manifest_path.read_text(encoding="utf-8"))
        except (OSError, ValueError) as exc:
            raise ValueError("overwrite is allowed only for a DocWen-owned document node") from exc
        if document.get("schema") != "docwen.document_node.v1" or document.get("node_name") != os.path.basename(root):
            raise ValueError("overwrite is allowed only for a matching DocWen-owned document node")

    def _unused_backup_path(self, root: str) -> str:
        counter = 0
        while True:
            counter += 1
            candidate = f"{root}.__docwen_backup_{counter:03d}"
            if not self._io_path(candidate).exists():
                return candidate

    @classmethod
    def _sha256_if_file(
        cls,
        path: str,
        cancellation: CancellationTokenView | None,
    ) -> str | None:
        if not path or not cls._io_path(path).is_file():
            return None
        return cls._file_integrity(path, cancellation)[1]

    @classmethod
    def _file_integrity(
        cls,
        path: str,
        cancellation: CancellationTokenView | None,
    ) -> tuple[int, str]:
        digest = hashlib.sha256()
        size = 0
        with cls._io_path(path).open("rb") as stream:
            while chunk := stream.read(_COPY_CHUNK_BYTES):
                cls._check_cancellation(cancellation)
                size += len(chunk)
                digest.update(chunk)
        return size, digest.hexdigest()

    def finalize_error(
        self,
        task_id: str,
        error_info,
        *,
        duration_ms: float = 0.0,
        input_bytes: int = 0,
    ) -> ConversionResult:
        """Create a failed ConversionResult without placing any artifacts."""
        from docwen_core.models.result import ConversionErrorInfo

        if isinstance(error_info, ConversionErrorInfo):
            err = error_info
        else:
            err = ConversionErrorInfo(
                error_type="conversion_failed",
                message=str(error_info),
            )

        return ConversionResult(
            task_id=task_id,
            success=False,
            artifacts=[],
            diagnostics=[],
            error=err,
            metrics=ConversionMetrics(
                duration_ms=duration_ms,
                input_bytes=input_bytes,
                output_bytes=0,
            ),
        )

    # ── Internal helpers ────────────────────────────────────────────

    @staticmethod
    def _check_cancellation(cancellation: CancellationTokenView | None) -> None:
        if cancellation is not None:
            cancellation.check()

    @classmethod
    def _acquire_thread_lock(
        cls,
        output_lock: Any,
        cancellation: CancellationTokenView | None,
    ) -> None:
        """Acquire the in-process lock without making cancellation wait forever."""
        while True:
            cls._check_cancellation(cancellation)
            if output_lock.acquire(timeout=_LOCK_POLL_SECONDS):
                return

    @classmethod
    def _finalization_lock_paths(
        cls,
        output_dir: str,
        artifacts: list[ArtifactManifest],
    ) -> tuple[str, ...]:
        """Return every output location whose contents this batch may mutate."""

        paths_by_key = {cls._lock_key(output_dir): output_dir}
        for artifact in artifacts:
            try:
                suggested = artifact.suggested_name or os.path.basename(artifact.staging_path)
                destination, _ = cls._safe_final_path(output_dir, suggested)
                parent = os.path.dirname(destination)
                paths_by_key.setdefault(cls._lock_key(parent), parent)
            except (OSError, ValueError):
                # The normal prepare path owns the typed per-artifact error.
                continue
        return tuple(paths_by_key[key] for key in sorted(paths_by_key))

    @classmethod
    @contextlib.contextmanager
    def _finalization_locks(
        cls,
        paths: tuple[str, ...],
        cancellation: CancellationTokenView | None,
    ):
        """Hold sorted in-process and OS locks for all concrete output parents."""

        acquired_thread_locks: list[Any] = []
        try:
            for path in paths:
                output_lock = cls._lock_for_output_dir(path)
                cls._acquire_thread_lock(output_lock, cancellation)
                acquired_thread_locks.append(output_lock)
            with contextlib.ExitStack() as stack:
                for path in paths:
                    stack.enter_context(cls._process_lock(path, cancellation))
                yield
        finally:
            for output_lock in reversed(acquired_thread_locks):
                output_lock.release()

    @classmethod
    @contextlib.contextmanager
    def _process_lock(
        cls,
        output_dir: str,
        cancellation: CancellationTokenView | None,
    ):
        """Hold a cancellation-aware OS lock keyed by resolved output directory."""
        resolved = cls._lock_key(output_dir)
        digest = hashlib.sha256(os.fsencode(resolved)).hexdigest()
        lock_dir = Path(tempfile.gettempdir()) / "docwen-output-finalizer-locks"
        lock_path = lock_dir / f"{digest}.lock"
        io_lock_path = cls._io_path(lock_path)
        io_lock_path.parent.mkdir(parents=True, exist_ok=True)
        lock_file = io_lock_path.open("a+b")
        acquired = False
        try:
            lock_file.seek(0, os.SEEK_END)
            if lock_file.tell() == 0:
                lock_file.write(b"\0")
                lock_file.flush()
            while not acquired:
                cls._check_cancellation(cancellation)
                try:
                    _try_lock_file(lock_file.fileno())
                    acquired = True
                except OSError as exc:
                    retryable = isinstance(exc, BlockingIOError) or exc.errno in {
                        errno.EACCES,
                        errno.EAGAIN,
                        errno.EDEADLK,
                    }
                    if not retryable:
                        raise
                    time.sleep(_LOCK_POLL_SECONDS)
            yield
        finally:
            if acquired:
                with contextlib.suppress(OSError):
                    _unlock_file(lock_file.fileno())
            lock_file.close()

    @staticmethod
    def _io_path(path: str | os.PathLike[str]) -> Path:
        """Return an absolute path suitable for internal filesystem I/O.

        Logical artifact paths stay prefix-free in manifests and diagnostics.
        On Windows only syscall operands crossing the legacy MAX_PATH boundary
        receive the extended-path prefix; this also covers UNC shares.
        """
        return filesystem_path(path)

    @staticmethod
    def _logical_io_spelling(path: str | os.PathLike[str]) -> str:
        """Remove only DocWen's internal Win32 extended-filesystem spelling."""

        raw = os.fsdecode(os.fspath(path))
        if sys.platform != "win32":
            return raw
        if raw.upper().startswith("\\\\?\\UNC\\"):
            return f"\\\\{raw[8:]}"
        if raw.startswith("\\\\?\\"):
            return raw[4:]
        return raw

    @classmethod
    def _public_exception_text(cls, exc: BaseException) -> str:
        """Render an exception without exposing an internal extended prefix."""

        if isinstance(exc, OSError) and any(
            value is not None
            for value in (
                getattr(exc, "winerror", None),
                exc.errno,
                exc.strerror,
                exc.filename,
                exc.filename2,
            )
        ):
            winerror = getattr(exc, "winerror", None)
            code = winerror if winerror is not None else exc.errno
            label = "WinError" if winerror is not None else "Errno"
            detail = exc.strerror or type(exc).__name__
            text = f"[{label} {code}] {detail}" if code is not None else detail
            if exc.filename is not None:
                first = cls._logical_io_spelling(exc.filename)
                text = f"{text}: {first!r}"
            if exc.filename2 is not None:
                second = cls._logical_io_spelling(exc.filename2)
                text = f"{text} -> {second!r}"
        else:
            text = str(exc)
        if sys.platform != "win32":
            return text
        marker = "\\\\?\\"
        unc_marker = "\\\\?\\UNC\\"
        while True:
            lowered = text.lower()
            index = lowered.find(marker.lower())
            if index < 0:
                return text
            if lowered.startswith(unc_marker.lower(), index):
                text = f"{text[:index]}\\\\{text[index + len(unc_marker) :]}"
            else:
                text = f"{text[:index]}{text[index + len(marker) :]}"

    @classmethod
    def _prepare_artifact(
        cls,
        artifact: ArtifactManifest,
        output_dir: str,
        overwrite_mode: str,
        input_path: str,
        cancellation: CancellationTokenView | None,
    ) -> _PreparedArtifact:
        """Copy an artifact to a private same-directory temp without publishing it."""
        if overwrite_mode not in {"error", "rename", "overwrite", "skip"}:
            raise ValueError(f"Unknown overwrite mode: {overwrite_mode!r}")

        reused = cls._reuse_identical_input_artifact(
            artifact,
            output_dir,
            overwrite_mode,
            input_path,
            cancellation,
        )
        suggested = artifact.suggested_name or os.path.basename(artifact.staging_path)
        destination, suggested = cls._safe_final_path(output_dir, suggested)
        rename_base = destination if overwrite_mode == "rename" else None
        if reused is not None:
            return _PreparedArtifact(
                artifact=artifact,
                suggested_name=suggested,
                destination=destination,
                rename_base=rename_base,
                reuse=reused[0],
            )

        io_destination = cls._io_path(destination)
        if io_destination.exists():
            if overwrite_mode != "rename" and not io_destination.is_file():
                raise IsADirectoryError(f"Existing output target is not a file: {destination}")
            if overwrite_mode == "skip":
                return _PreparedArtifact(
                    artifact=artifact,
                    suggested_name=suggested,
                    destination=destination,
                    rename_base=rename_base,
                    skip_existing=True,
                )
            if overwrite_mode == "rename":
                destination = cls._rename_path(destination)
            elif overwrite_mode == "error":
                raise FileExistsError(f"Output target already exists: {destination}")

        if not cls._io_path(artifact.staging_path).is_file():
            raise FileNotFoundError(f"Staging artifact is not a file: {artifact.staging_path}")

        parent = os.path.dirname(destination)
        io_parent = filesystem_path(parent, force_extended=sys.platform == "win32")
        io_parent.mkdir(parents=True, exist_ok=True)
        cls._ensure_contained(output_dir, destination)
        # ``mkstemp`` appends its private random name.  A parent that is still
        # below MAX_PATH can therefore produce a child above it, so force the
        # filesystem namespace for this append operation on Windows.
        file_descriptor, temp_path = tempfile.mkstemp(prefix=_TEMP_PREFIX, dir=os.fspath(io_parent))
        os.close(file_descriptor)
        try:
            cls._copy_to_temp(artifact.staging_path, temp_path, cancellation)
        except BaseException:
            with contextlib.suppress(OSError):
                cls._io_path(temp_path).unlink()
            raise
        return _PreparedArtifact(
            artifact=artifact,
            suggested_name=suggested,
            destination=destination,
            rename_base=rename_base,
            temp_path=temp_path,
        )

    @classmethod
    def _copy_to_temp(
        cls,
        source: str,
        temp_path: str,
        cancellation: CancellationTokenView | None,
    ) -> None:
        """Copy into a private file, checking cancellation between chunks."""
        io_source = OutputFinalizer._io_path(source)
        io_temp = OutputFinalizer._io_path(temp_path)
        with io_source.open("rb") as source_file, io_temp.open("wb") as temp_file:
            while True:
                cls._check_cancellation(cancellation)
                chunk = source_file.read(_COPY_CHUNK_BYTES)
                if not chunk:
                    break
                temp_file.write(chunk)
            temp_file.flush()
            os.fsync(temp_file.fileno())
        shutil.copystat(io_source, io_temp)

    @classmethod
    def _commit_prepared(
        cls,
        item: _PreparedArtifact,
        output_dir: str,
        overwrite_mode: str,
    ) -> tuple[ArtifactManifest, int]:
        """Publish one prepared artifact atomically without cancellation checks."""
        cls._ensure_contained(output_dir, item.destination)
        if item.reuse is not None:
            return item.reuse, 0
        if item.skip_existing:
            if not cls._io_path(item.destination).is_file():
                raise FileNotFoundError(f"Existing output target disappeared: {item.destination}")
            return cls._placed_manifest(item, item.destination, skipped=True), 0
        if item.temp_path is None:
            raise RuntimeError("Prepared artifact has no commit source")

        destination = item.destination
        if overwrite_mode == "overwrite":
            os.replace(cls._io_path(item.temp_path), cls._io_path(destination))
        else:
            while True:
                try:
                    cls._publish_no_clobber(item.temp_path, destination)
                    break
                except FileExistsError:
                    if overwrite_mode != "rename":
                        raise
                    destination = cls._rename_path(item.rename_base or destination)
                    cls._ensure_contained(output_dir, destination)
        if not cls._io_path(item.temp_path).exists():
            item.temp_path = None
        return cls._placed_manifest(item, destination), cls._io_path(destination).stat().st_size

    @classmethod
    def _publish_no_clobber(cls, temp_path: str, destination: str) -> None:
        """Atomically publish a new path without replacing an external winner."""
        io_temp = cls._io_path(temp_path)
        io_destination = cls._io_path(destination)
        if sys.platform == "win32":
            # Windows rename is no-replace and works on filesystems that do not
            # support hard links (for example FAT-family removable media).
            os.rename(io_temp, io_destination)
            return
        os.link(io_temp, io_destination)
        with contextlib.suppress(OSError):
            io_temp.unlink()

    @staticmethod
    def _placed_manifest(
        item: _PreparedArtifact,
        destination: str,
        *,
        skipped: bool = False,
    ) -> ArtifactManifest:
        metadata = dict(item.artifact.metadata)
        if skipped:
            metadata.update({"skipped": True, "reason": "file_exists"})
        return ArtifactManifest(
            artifact_id=item.artifact.artifact_id,
            kind=item.artifact.kind,
            staging_path=destination,
            suggested_name=item.suggested_name,
            media_type=item.artifact.media_type,
            metadata=metadata,
            is_primary=item.artifact.is_primary,
        )

    @classmethod
    def _cleanup_stale_temps(cls, parent: str) -> None:
        """Remove only aged members of the reserved pre-commit temp family."""
        now = time.time()
        with os.scandir(cls._io_path(parent)) as entries:
            for entry in entries:
                if not entry.name.startswith(_TEMP_PREFIX) or not entry.is_file(follow_symlinks=False):
                    continue
                try:
                    inspected = entry.stat(follow_symlinks=False)
                    age_seconds = now - inspected.st_ctime
                except FileNotFoundError:
                    continue
                if age_seconds < _STALE_TEMP_MIN_AGE_SECONDS:
                    continue
                target = cls._io_path(entry.path)
                try:
                    current = target.stat(follow_symlinks=False)
                except FileNotFoundError:
                    continue
                inspected_identity = (
                    inspected.st_ctime_ns,
                    inspected.st_mtime_ns,
                    inspected.st_size,
                    inspected.st_mode,
                )
                current_identity = (
                    current.st_ctime_ns,
                    current.st_mtime_ns,
                    current.st_size,
                    current.st_mode,
                )
                if inspected.st_ino and current.st_ino:
                    inspected_identity += (inspected.st_dev, inspected.st_ino)
                    current_identity += (current.st_dev, current.st_ino)
                if current_identity != inspected_identity:
                    continue
                with contextlib.suppress(FileNotFoundError):
                    target.unlink()

    @staticmethod
    def _reuse_identical_input_artifact(
        artifact: ArtifactManifest,
        output_dir: str,
        overwrite_mode: str,
        input_path: str,
        cancellation: CancellationTokenView | None = None,
    ) -> tuple[ArtifactManifest, int] | None:
        """Reuse an unchanged retained input already in the final directory.

        Image-to-Markdown file mode registers a staging copy of the input as a
        non-primary retained artifact.  When output remains beside the input,
        that input already occupies the retained artifact's suggested path.
        Treating it as an external collision would either fail a valid request,
        overwrite the source, or orphan the Markdown link.  Reuse the exact
        input under every overwrite policy, but only when the staging bytes
        prove it is the same retained file.
        """
        if artifact.is_primary or not input_path:
            return None

        final_path, suggested = OutputFinalizer._safe_final_path(
            output_dir,
            artifact.suggested_name or os.path.basename(artifact.staging_path),
        )
        input_abs = os.path.abspath(input_path)
        if os.path.normcase(final_path) != os.path.normcase(input_abs):
            return None
        if (
            not OutputFinalizer._io_path(artifact.staging_path).is_file()
            or not OutputFinalizer._io_path(input_abs).is_file()
        ):
            return None
        if not OutputFinalizer._files_identical(artifact.staging_path, input_abs, cancellation):
            raise ValueError("Retained artifact collides with its input path but has different bytes")

        reused = ArtifactManifest(
            artifact_id=artifact.artifact_id,
            kind=artifact.kind,
            staging_path=input_abs,
            suggested_name=suggested,
            media_type=artifact.media_type,
            metadata={**artifact.metadata, "reused_input": True},
            is_primary=artifact.is_primary,
        )
        return reused, 0

    @staticmethod
    def _files_identical(
        first_path: str,
        second_path: str,
        cancellation: CancellationTokenView | None = None,
    ) -> bool:
        """Compare two files without loading an unbounded artifact into memory."""
        first_io = OutputFinalizer._io_path(first_path)
        second_io = OutputFinalizer._io_path(second_path)
        if first_io.stat().st_size != second_io.stat().st_size:
            return False
        with first_io.open("rb") as first, second_io.open("rb") as second:
            while True:
                OutputFinalizer._check_cancellation(cancellation)
                first_chunk = first.read(1024 * 1024)
                second_chunk = second.read(1024 * 1024)
                if first_chunk != second_chunk:
                    return False
                if not first_chunk:
                    return True

    @classmethod
    def _lock_for_output_dir(cls, output_dir: str) -> Any:
        """Return a process-wide lock shared by finalizers targeting one directory."""
        key = cls._lock_key(output_dir)
        with cls._output_locks_guard:
            output_lock = cls._output_locks.get(key)
            if output_lock is None:
                output_lock = threading.RLock()
                cls._output_locks[key] = output_lock
            return output_lock

    @classmethod
    def _lock_key(cls, path: str) -> str:
        """Return one prefix-free identity key for a concrete output location."""

        absolute = os.path.abspath(path)
        io_path = filesystem_path(absolute, force_extended=sys.platform == "win32")
        resolved = os.path.realpath(io_path)
        return os.path.normcase(cls._logical_io_spelling(resolved))

    @staticmethod
    def _resolve_output_dir(policy: OutputPolicy, input_path: str) -> str:
        """Determine the final output directory."""
        if policy.output_path:
            if policy.output_dir:
                raise ValueError("output_path and output_dir are mutually exclusive")
            if policy.date_subfolder:
                raise ValueError("output_path cannot be combined with date_subfolder")
            output_path = os.path.abspath(policy.output_path)
            if not os.path.basename(output_path):
                raise ValueError("output_path must name a file")
            base = os.path.dirname(output_path) or "."
        elif policy.output_dir:
            base = policy.output_dir
        elif input_path:
            base = os.path.dirname(input_path) or "."
        else:
            base = "."

        if policy.date_subfolder:
            from datetime import date

            today = date.today()
            if policy.date_subfolder == "iso":
                sub = today.isoformat()
            elif policy.date_subfolder == "compact":
                sub = today.strftime("%Y%m%d")
            elif policy.date_subfolder == "chinese":
                sub = f"{today.year}年{today.month:02d}月{today.day:02d}日"
            else:
                sub = policy.date_subfolder
            return os.path.join(base, sub)

        return base

    @staticmethod
    def _artifacts_for_policy(
        artifacts: list[ArtifactManifest],
        policy: OutputPolicy,
    ) -> list[ArtifactManifest]:
        """Apply an exact primary output name without changing auxiliary names."""

        if not policy.output_path:
            return artifacts
        primary_indexes = [index for index, artifact in enumerate(artifacts) if artifact.is_primary]
        if len(primary_indexes) != 1:
            raise ValueError("output_path requires exactly one primary artifact")
        output_name = os.path.basename(os.path.abspath(policy.output_path))
        if not output_name:
            raise ValueError("output_path must name a file")
        projected = list(artifacts)
        index = primary_indexes[0]
        projected[index] = replace(projected[index], suggested_name=output_name)
        return projected

    @staticmethod
    def _safe_final_path(output_dir: str, suggested_name: str) -> tuple[str, str]:
        """Resolve a plugin-suggested relative name inside ``output_dir``."""
        suggested = os.path.normpath(suggested_name)
        if (
            not suggested
            or suggested == "."
            or os.path.isabs(suggested)
            or os.path.splitdrive(suggested)[0]
            or suggested.startswith("..")
            or f"{os.pardir}{os.sep}" in suggested
        ):
            raise ValueError(f"Unsafe artifact suggested_name: {suggested_name!r}")

        output_abs = os.path.abspath(output_dir)
        final_path = os.path.abspath(os.path.join(output_abs, suggested))
        if os.path.commonpath([output_abs, final_path]) != output_abs:
            raise ValueError(f"Unsafe artifact suggested_name: {suggested_name!r}")
        OutputFinalizer._ensure_contained(output_abs, final_path)
        return final_path, suggested

    @staticmethod
    def _ensure_contained(output_dir: str, final_path: str) -> None:
        """Reject descendant symlink/junction resolution outside the selected root."""
        output_real = OutputFinalizer._logical_io_spelling(
            os.path.realpath(filesystem_path(os.path.abspath(output_dir), force_extended=sys.platform == "win32"))
        )
        final_real = OutputFinalizer._logical_io_spelling(
            os.path.realpath(filesystem_path(os.path.abspath(final_path), force_extended=sys.platform == "win32"))
        )
        try:
            common = os.path.commonpath([output_real, final_real])
        except ValueError as exc:
            raise ValueError(f"Resolved artifact path escapes output directory: {final_path!r}") from exc
        if os.path.normcase(common) != os.path.normcase(output_real):
            raise ValueError(f"Resolved artifact path escapes output directory: {final_path!r}")

    @staticmethod
    def _rename_path(path: str) -> str:
        """Generate a unique path using the stable numbered-suffix contract."""
        if not OutputFinalizer._io_path(path).exists():
            return path
        base, ext = os.path.splitext(path)
        counter = 1
        while True:
            candidate = f"{base}_{counter:03d}{ext}"
            if not OutputFinalizer._io_path(candidate).exists():
                return candidate
            counter += 1
