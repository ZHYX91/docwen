"""OutputFinalizer — reads staging artifacts and places them at the final output path.

This is the ONLY component allowed to write to the final output directory.
Plugins MUST NOT call any method on this class directly.
"""

from __future__ import annotations

import contextlib
import errno
import hashlib
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
    DocumentNodeLayoutPlan,
    has_markdown_artifacts,
    plan_document_node_layout,
    relocated_markdown_bytes,
)
from docwen_runtime.output.node_audit import stage_node_audit
from docwen_runtime.path_io import filesystem_path

if TYPE_CHECKING:
    from docwen_core.models.document_node import ConversionIdentity
    from docwen_core.models.request import OutputPolicy
    from docwen_core.protocols.execution_context import CancellationTokenView
    from docwen_runtime.output.manifest import OutputManifestDocument


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
        group_outputs: bool = False,
        identity: ConversionIdentity | None = None,
        audit_document: OutputManifestDocument | None = None,
    ) -> ConversionResult:
        """Finalize a set of artifacts."""
        self._check_cancellation(cancellation)
        output_dir = self._resolve_output_dir(policy, input_path)
        node_plan: DocumentNodeLayoutPlan | None = None
        if artifacts and (group_outputs or has_markdown_artifacts(artifacts)):
            in_place_markdown = bool(
                policy.output_path
                and input_path
                and os.path.normcase(os.path.abspath(policy.output_path))
                == os.path.normcase(os.path.abspath(input_path))
            )
            if policy.output_path and not in_place_markdown:
                raise ValueError("Grouped conversion output requires an output parent directory, not output_path")
            if in_place_markdown:
                artifacts = self._artifacts_for_policy(artifacts, policy)
            else:
                node_plan = plan_document_node_layout(
                    task_id=task_id,
                    artifacts=artifacts,
                    input_path=input_path,
                    identity=identity,
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
                    audit_document=audit_document,
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
        filesystem_path(output_dir, force_extended=sys.platform == "win32").mkdir(parents=True, exist_ok=True)

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
                            message=f"Failed to place artifact {artifact.artifact_id!r}: {self._public_exception_text(exc)}",
                            code="FINALIZER_PLACE_ERROR",
                        )
                    )

            self._check_cancellation(cancellation)
            for item in prepared:
                try:
                    final_artifact, written_bytes = self._commit_prepared(item, output_dir, policy.overwrite_mode)
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
                    try:
                        self._io_path(item.temp_path).unlink()
                    except FileNotFoundError:
                        pass
                    except OSError as exc:
                        diagnostics.append(
                            ConversionDiagnostic(
                                level="warning",
                                message=f"Temporary output cleanup failed: {self._public_exception_text(exc)}",
                                code="FINALIZER_CLEANUP_FAILED",
                            )
                        )

        attempted_artifacts = len(artifacts)
        placed_artifacts = len(final_artifacts)
        error: ConversionErrorInfo | None = None
        metrics_extra: dict[str, Any] = {"output_dir": output_dir}

        if attempted_artifacts == 0:
            summary_code = "FINALIZER_NO_ARTIFACTS"
            summary_message = "No output artifacts were provided for finalization"
            diagnostics.append(ConversionDiagnostic(level="error", message=summary_message, code=summary_code))
            error = ConversionErrorInfo(
                error_type="output_finalization_failed",
                message=summary_message,
                diagnostic_code=summary_code,
            )
        elif failed_artifacts:
            summary_code = "FINALIZER_PARTIAL" if placed_artifacts else "FINALIZER_FAILED"
            summary_message = f"Placed {placed_artifacts} of {attempted_artifacts} artifact(s) in {output_dir}"
            diagnostics.append(ConversionDiagnostic(level="error", message=summary_message, code=summary_code))
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
        audit_document: OutputManifestDocument | None = None,
    ) -> ConversionResult:
        """Publish business artifacts and an explicitly requested audit atomically.

        Integrity remains in memory. No sidecar or persistent ownership index
        is needed, and arbitrary existing directories are never replaced.
        """
        self._check_cancellation(cancellation)
        if policy.overwrite_mode not in {"error", "rename", "skip"}:
            return self._document_node_failure(
                task_id,
                "Result directories support rename, error, or identical-content skip. "
                "Overwrite applies only to individual files; choose a new output parent instead.",
                duration_ms=duration_ms,
                input_bytes=input_bytes,
                output_dir=output_dir,
                root_name=plan.root_name,
                code="DOCUMENT_NODE_POLICY_UNSUPPORTED",
            )
        output_io = self._io_path(output_dir)
        output_io.mkdir(parents=True, exist_ok=True)
        selected = plan
        collision = 0
        while os.path.lexists(self._io_path(os.path.join(output_dir, selected.root_name))):
            if policy.overwrite_mode == "rename":
                collision += 1
                selected = plan.rebase_root(plan.identity.node_name(collision=collision))
                continue
            if policy.overwrite_mode == "skip":
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
        temp_root = ""
        diagnostics: list[ConversionDiagnostic] = []
        placed: list[ArtifactManifest] = []
        output_bytes = 0
        reused = False
        failure: ConversionResult | None = None
        try:
            self._ensure_contained(output_dir, final_root)
            temp_parent = filesystem_path(output_dir, force_extended=sys.platform == "win32")
            temp_root = tempfile.mkdtemp(prefix=".__docwen-node-", dir=os.fspath(temp_parent))
            self._ensure_contained(output_dir, temp_root)
            prepared: list[ArtifactManifest] = []
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
                    with destination_io.open("xb") as stream:
                        stream.write(payload)
                        stream.flush()
                        os.fsync(stream.fileno())
                    with contextlib.suppress(OSError):
                        shutil.copystat(self._io_path(artifact.staging_path), destination_io)
                else:
                    self._copy_to_temp(artifact.staging_path, destination, cancellation)
                size_bytes, sha256 = self._file_integrity(destination, cancellation)
                prepared.append(replace(artifact, size_bytes=size_bytes, sha256=sha256))
                output_bytes += size_bytes

            if audit_document is not None:
                audit = stage_node_audit(selected, temp_root, audit_document)
                size_bytes, sha256 = self._file_integrity(audit.staging_path, cancellation)
                prepared.append(replace(audit, size_bytes=size_bytes, sha256=sha256))
                output_bytes += size_bytes

            self._check_cancellation(cancellation)
            if policy.overwrite_mode == "skip" and os.path.lexists(self._io_path(final_root)):
                self._verify_identical_node(final_root, temp_root, prepared, cancellation)
                reused = True
                output_bytes = 0
            else:
                # Ordinary POSIX rename may replace an empty directory created
                # by another program after our collision check. Never clobber it.
                self._publish_directory_no_clobber(temp_root, final_root)
                temp_root = ""
            placed = [
                replace(
                    artifact,
                    staging_path=os.path.join(final_root, *artifact.logical_path.split("/")[1:]),
                    metadata={
                        **artifact.metadata,
                        "document_node_committed": not reused,
                        "document_node_reused": reused,
                    },
                )
                for artifact in prepared
                if artifact.logical_path is not None
            ]
            diagnostics.append(
                ConversionDiagnostic(
                    level="info",
                    message=f"{'Reused identical' if reused else 'Published'} document node {final_root}",
                    code="DOCUMENT_NODE_REUSED" if reused else "FINALIZER_DONE",
                )
            )
        except CancellationRequested:
            raise
        except Exception as exc:
            failure = self._document_node_failure(
                task_id,
                self._public_exception_text(exc),
                duration_ms=duration_ms,
                input_bytes=input_bytes,
                output_dir=output_dir,
                root_name=selected.root_name,
                code="DOCUMENT_NODE_SKIP_MISMATCH"
                if policy.overwrite_mode == "skip"
                else "DOCUMENT_NODE_PUBLISH_FAILED",
            )
        finally:
            if temp_root and self._io_path(temp_root).exists():
                try:
                    shutil.rmtree(self._io_path(temp_root))
                except OSError as exc:
                    diagnostics.append(
                        ConversionDiagnostic(
                            level="warning",
                            message=f"Temporary output cleanup failed: {self._public_exception_text(exc)}",
                            code="FINALIZER_CLEANUP_FAILED",
                        )
                    )
        if failure is not None:
            failure.diagnostics.extend(diagnostics)
            return failure
        return ConversionResult(
            task_id=task_id,
            success=True,
            artifacts=placed,
            diagnostics=diagnostics,
            metrics=ConversionMetrics(
                duration_ms=duration_ms,
                input_bytes=input_bytes,
                output_bytes=output_bytes,
                extra={
                    "output_dir": output_dir,
                    "document_node_root": final_root,
                    "document_node_schema": "docwen.document_node.v1",
                    "document_node_reused": reused,
                },
            ),
        )

    @classmethod
    def _verify_identical_node(
        cls,
        root: str,
        prepared_root: str,
        artifacts: list[ArtifactManifest],
        cancellation: CancellationTokenView | None,
    ) -> None:
        """Reuse only a complete byte-identical tree; never trust its filename."""

        def tree(path: str) -> set[str]:
            native = cls._io_path(path)
            if not native.is_dir() or native.is_symlink() or native.is_junction():
                raise ValueError("existing result must be a regular directory")
            entries: set[str] = set()
            for directory, dirs, files in os.walk(native, followlinks=False):
                cls._check_cancellation(cancellation)
                for name in (*dirs, *files):
                    item = Path(directory) / name
                    if item.is_symlink() or item.is_junction():
                        raise ValueError("existing result contains a link or junction")
                    relative = item.relative_to(native).as_posix()
                    entries.add(relative + ("/" if item.is_dir() else ""))
            return entries

        if tree(root) != tree(prepared_root):
            raise ValueError("existing result contains missing or extra output paths")
        for artifact in artifacts:
            if artifact.logical_path is None:
                raise ValueError("prepared artifact has no logical path")
            path = os.path.join(root, *artifact.logical_path.split("/")[1:])
            if cls._file_integrity(path, cancellation) != (artifact.size_bytes, artifact.sha256):
                raise ValueError(f"existing result differs from prepared output: {artifact.logical_path}")

    @classmethod
    def _publish_directory_no_clobber(cls, source: str, destination: str) -> None:
        if sys.platform in {"win32", "linux"}:
            cls._publish_no_clobber(source, destination)
            return
        if sys.platform == "darwin":
            import ctypes

            libc = ctypes.CDLL(None, use_errno=True)
            rename_exclusive = libc.renamex_np
            rename_exclusive.argtypes = [ctypes.c_char_p, ctypes.c_char_p, ctypes.c_uint]
            rename_exclusive.restype = ctypes.c_int
            if rename_exclusive(os.fsencode(source), os.fsencode(destination), 4) != 0:  # RENAME_EXCL
                error = ctypes.get_errno()
                raise OSError(error, os.strerror(error), destination)
            return
        raise OSError(errno.ENOSYS, "Atomic no-replace directory rename is unavailable", destination)

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
        from docwen_core.models.result import ConversionErrorInfo

        if isinstance(error_info, ConversionErrorInfo):
            err = error_info
        else:
            err = ConversionErrorInfo(error_type="conversion_failed", message=str(error_info))

        return ConversionResult(
            task_id=task_id,
            success=False,
            artifacts=[],
            diagnostics=[],
            error=err,
            metrics=ConversionMetrics(duration_ms=duration_ms, input_bytes=input_bytes, output_bytes=0),
        )

    @staticmethod
    def _check_cancellation(cancellation: CancellationTokenView | None) -> None:
        if cancellation is not None:
            cancellation.check()

    @classmethod
    def _acquire_thread_lock(cls, output_lock: Any, cancellation: CancellationTokenView | None) -> None:
        while True:
            cls._check_cancellation(cancellation)
            if output_lock.acquire(timeout=_LOCK_POLL_SECONDS):
                return

    @classmethod
    def _finalization_lock_paths(cls, output_dir: str, artifacts: list[ArtifactManifest]) -> tuple[str, ...]:
        paths_by_key = {cls._lock_key(output_dir): output_dir}
        for artifact in artifacts:
            try:
                suggested = artifact.suggested_name or os.path.basename(artifact.staging_path)
                destination, _ = cls._safe_final_path(output_dir, suggested)
                parent = os.path.dirname(destination)
                paths_by_key.setdefault(cls._lock_key(parent), parent)
            except (OSError, ValueError):
                continue
        return tuple(paths_by_key[key] for key in sorted(paths_by_key))

    @classmethod
    @contextlib.contextmanager
    def _finalization_locks(cls, paths: tuple[str, ...], cancellation: CancellationTokenView | None):
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
    def _process_lock(cls, output_dir: str, cancellation: CancellationTokenView | None):
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
        return filesystem_path(path)

    @staticmethod
    def _logical_io_spelling(path: str | os.PathLike[str]) -> str:
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
        if overwrite_mode not in {"error", "rename", "overwrite", "skip"}:
            raise ValueError(f"Unknown overwrite mode: {overwrite_mode!r}")

        reused = cls._reuse_identical_input_artifact(artifact, output_dir, overwrite_mode, input_path, cancellation)
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
        file_descriptor, temp_path = tempfile.mkstemp(prefix=_TEMP_PREFIX, dir=os.fspath(io_parent))
        os.close(file_descriptor)
        try:
            cls._copy_to_temp(artifact.staging_path, temp_path, cancellation)
            size_bytes, sha256 = cls._file_integrity(temp_path, cancellation)
        except BaseException:
            with contextlib.suppress(OSError):
                cls._io_path(temp_path).unlink()
            raise
        return _PreparedArtifact(
            artifact=replace(artifact, size_bytes=size_bytes, sha256=sha256),
            suggested_name=suggested,
            destination=destination,
            rename_base=rename_base,
            temp_path=temp_path,
        )

    @classmethod
    def _copy_to_temp(cls, source: str, temp_path: str, cancellation: CancellationTokenView | None) -> None:
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
        cls, item: _PreparedArtifact, output_dir: str, overwrite_mode: str
    ) -> tuple[ArtifactManifest, int]:
        cls._ensure_contained(output_dir, item.destination)
        if item.reuse is not None:
            return item.reuse, 0
        if item.skip_existing:
            if not cls._io_path(item.destination).is_file():
                raise FileNotFoundError(f"Existing output target disappeared: {item.destination}")
            size_bytes, sha256 = cls._file_integrity(item.destination, None)
            item.artifact = replace(item.artifact, size_bytes=size_bytes, sha256=sha256)
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
        return cls._placed_manifest(item, destination), item.artifact.size_bytes or 0

    @classmethod
    def _publish_no_clobber(cls, temp_path: str, destination: str) -> None:
        io_temp = cls._io_path(temp_path)
        io_destination = cls._io_path(destination)
        if sys.platform == "win32":
            os.rename(io_temp, io_destination)
            return
        if sys.platform == "linux":
            import ctypes
            import errno

            libc = ctypes.CDLL(None, use_errno=True)
            rename_no_replace = getattr(libc, "renameat2", None)
            if rename_no_replace is None:
                raise OSError(errno.ENOSYS, "Atomic no-replace rename is unavailable", str(io_destination))
            rename_no_replace.argtypes = [ctypes.c_int, ctypes.c_char_p, ctypes.c_int, ctypes.c_char_p, ctypes.c_uint]
            rename_no_replace.restype = ctypes.c_int
            if rename_no_replace(-100, os.fsencode(io_temp), -100, os.fsencode(io_destination), 1) != 0:
                error = ctypes.get_errno()
                raise OSError(error, os.strerror(error), str(io_destination))
            return
        os.link(io_temp, io_destination)
        with contextlib.suppress(OSError):
            io_temp.unlink()

    @staticmethod
    def _placed_manifest(item: _PreparedArtifact, destination: str, *, skipped: bool = False) -> ArtifactManifest:
        metadata = dict(item.artifact.metadata)
        if skipped:
            metadata.update({"skipped": True, "reason": "file_exists"})
        return replace(
            item.artifact,
            staging_path=destination,
            suggested_name=item.suggested_name,
            metadata=metadata,
        )

    @classmethod
    def _cleanup_stale_temps(cls, parent: str) -> None:
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

        size_bytes, sha256 = OutputFinalizer._file_integrity(input_abs, cancellation)
        reused = replace(
            artifact,
            staging_path=input_abs,
            suggested_name=suggested,
            metadata={**artifact.metadata, "reused_input": True},
            size_bytes=size_bytes,
            sha256=sha256,
        )
        return reused, 0

    @staticmethod
    def _files_identical(
        first_path: str,
        second_path: str,
        cancellation: CancellationTokenView | None = None,
    ) -> bool:
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
        key = cls._lock_key(output_dir)
        with cls._output_locks_guard:
            output_lock = cls._output_locks.get(key)
            if output_lock is None:
                output_lock = threading.RLock()
                cls._output_locks[key] = output_lock
            return output_lock

    @classmethod
    def _lock_key(cls, path: str) -> str:
        absolute = os.path.abspath(path)
        io_path = filesystem_path(absolute, force_extended=sys.platform == "win32")
        resolved = os.path.realpath(io_path)
        return os.path.normcase(cls._logical_io_spelling(resolved))

    @staticmethod
    def _resolve_output_dir(policy: OutputPolicy, input_path: str) -> str:
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
    def _artifacts_for_policy(artifacts: list[ArtifactManifest], policy: OutputPolicy) -> list[ArtifactManifest]:
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
        if not OutputFinalizer._io_path(path).exists():
            return path
        base, ext = os.path.splitext(path)
        counter = 1
        while True:
            candidate = f"{base}_{counter:03d}{ext}"
            if not OutputFinalizer._io_path(candidate).exists():
                return candidate
            counter += 1
