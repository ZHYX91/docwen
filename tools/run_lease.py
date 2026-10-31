"""Canonical identities and states for engineering run leases."""

from __future__ import annotations

import json
import os
import shutil
import stat
import sys
import tempfile
from collections.abc import Iterator
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import Any

from tools.process_identity import observe

STATES = frozenset(
    {
        "active",
        "completed-success",
        "retained-failure",
        "retained-interrupted",
        "retained-cleanup-failure",
        "retained-manual",
    }
)


@dataclass
class ManagedRun:
    root: Path
    state: str = "completed-success"


def _same_entry(path: Path, expected: os.stat_result) -> bool:
    actual = path.lstat()
    return (
        (actual.st_dev, actual.st_ino) == (expected.st_dev, expected.st_ino)
        and not stat.S_ISLNK(actual.st_mode)
        and not (getattr(actual, "st_file_attributes", 0) & getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0))
    )


@dataclass
class OwnedRunLease:
    """A captured lease identity, never authority to adopt an unknown run."""

    root: Path
    root_identity: os.stat_result
    marker_identity: os.stat_result
    lease: dict[str, Any]
    owner: str
    marker_content: bytes

    @property
    def marker(self) -> Path:
        return self.root / ".docwen-temp-lease.json"

    @classmethod
    def capture(cls, root: Path, *, owner: str, kind: str) -> OwnedRunLease:
        root_identity = root.lstat()
        marker = root / ".docwen-temp-lease.json"
        marker_identity = marker.lstat()
        if not stat.S_ISREG(marker_identity.st_mode) or not _same_entry(marker, marker_identity):
            raise ValueError(f"managed_run_marker_changed:{marker}")
        with marker.open("rb") as stream:
            opened = os.fstat(stream.fileno())
            if (opened.st_dev, opened.st_ino) != (marker_identity.st_dev, marker_identity.st_ino):
                raise ValueError(f"managed_run_marker_changed:{marker}")
            marker_content = stream.read()
            lease = json.loads(marker_content)
        if (
            not isinstance(lease, dict)
            or lease.get("schemaVersion") != 1
            or lease.get("state") not in STATES
            or lease.get("owner") != owner
            or lease.get("kind") != kind
            or lease.get("root") != str(root)
            or not _same_entry(marker, marker_identity)
        ):
            raise ValueError(f"managed_run_lease_identity_changed:{root}")
        captured = cls(root, root_identity, marker_identity, lease, owner, marker_content)
        captured.check_root()
        return captured

    def check_root(self) -> None:
        if self.root.resolve(strict=True) != self.root or not _same_entry(self.root, self.root_identity):
            raise ValueError(f"managed_run_identity_changed:{self.root}")

    @contextmanager
    def pin_root(self) -> Iterator[None]:
        """Keep the verified directory object alive through cleanup and recovery.

        The pin prevents reuse of a deleted root's inode/file ID during this
        interval. It does not turn path replacement into an atomic CAS.
        Windows permits delete sharing so ordinary recursive cleanup can finish;
        its final directory removal completes when this handle is closed.
        """
        self.check_root()
        if os.name == "nt":
            try:
                import msvcrt

                import pywintypes
                import win32api
                import win32file
            except ImportError as error:
                raise OSError("managed_run_root_pin_unavailable:windows") from error
            name = str(self.root)
            if not name.startswith("\\\\?\\"):
                name = "\\\\?\\UNC\\" + name[2:] if name.startswith("\\\\") else "\\\\?\\" + name
            try:
                handle = win32file.CreateFile(
                    name, 0x80, 7, None, 3, 0x00200000 | 0x02000000, None
                )  # READ_ATTRIBUTES; share READ/WRITE/DELETE; no reparse traversal
            except pywintypes.error as error:
                raise OSError(f"managed_run_root_pin_unavailable:{self.root}: {error}") from error
            raw_handle = handle.Detach()
            try:
                descriptor = msvcrt.open_osfhandle(raw_handle, os.O_RDONLY | os.O_BINARY)
            except BaseException as primary:
                try:
                    win32api.CloseHandle(raw_handle)
                except pywintypes.error as secondary:
                    self.report_secondary(primary, secondary)
                raise
        elif os.name == "posix":
            directory = getattr(os, "O_DIRECTORY", None)
            nofollow = getattr(os, "O_NOFOLLOW", None)
            if directory is None or nofollow is None:
                raise OSError("managed_run_root_pin_unavailable:directory_flags")
            descriptor = os.open(self.root, os.O_RDONLY | directory | nofollow | getattr(os, "O_CLOEXEC", 0))
        else:
            raise OSError("managed_run_root_pin_unavailable:platform")
        primary: BaseException | None = None
        try:
            opened = os.fstat(descriptor)
            if (
                not stat.S_ISDIR(opened.st_mode)
                or not opened.st_ino
                or getattr(opened, "st_file_attributes", 0) & getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0)
                or (opened.st_dev, opened.st_ino) != (self.root_identity.st_dev, self.root_identity.st_ino)
            ):
                raise ValueError(f"managed_run_identity_changed:{self.root}")
            self.check_root()
            yield
        except BaseException as error:
            primary = error
            raise
        finally:
            try:
                os.close(descriptor)
            except OSError as secondary:
                if primary is None:
                    raise
                self.report_secondary(primary, secondary)

    def report_secondary(self, primary: BaseException, secondary: BaseException) -> None:
        note = f"Managed run closeout failed; inspect {self.root}: {secondary}"
        primary.add_note(note)
        print(note, file=sys.stderr)

    def save(self, state: str, *, restore_missing: bool = False) -> None:
        self.check_root()
        try:
            marker_matches = (
                _same_entry(self.marker, self.marker_identity) and self.marker.read_bytes() == self.marker_content
            )
        except FileNotFoundError:
            if not restore_missing:
                raise
            # Partial recursive cleanup may have removed our marker. The root
            # still has its original identity; exclusive creation cannot replace
            # a new owner's marker that appears after this check.
            transition(self.lease, root=self.root, owner=self.owner, state=state)
            with self.marker.open("x", encoding="utf-8") as stream:
                self.marker_identity = os.fstat(stream.fileno())
                stream.write(json.dumps(self.lease))
                stream.flush()
                os.fsync(stream.fileno())
            self.marker_content = json.dumps(self.lease).encode("utf-8")
            return
        if not marker_matches:
            raise ValueError(f"managed_run_marker_changed:{self.marker}")
        transition(self.lease, root=self.root, owner=self.owner, state=state)
        temporary = self.root / ".docwen-temp-lease.next"
        with temporary.open("x", encoding="utf-8") as stream:
            next_identity = os.fstat(stream.fileno())
            json.dump(self.lease, stream)
            stream.flush()
            os.fsync(stream.fileno())
        self.check_root()
        if not _same_entry(self.marker, self.marker_identity) or self.marker.read_bytes() != self.marker_content:
            raise ValueError(f"managed_run_marker_changed:{self.marker}")
        temporary.replace(self.marker)
        self.marker_identity = next_identity
        self.marker_content = json.dumps(self.lease).encode("utf-8")

    def record_recovery(
        self,
        primary: BaseException,
        phase: str,
        *,
        outcome: str,
        recording_error: BaseException,
        execution_error: BaseException | None = None,
    ) -> None:
        recovery = self.root.parent / f"{self.root.name}.recovery.json"
        try:
            # This compact recovery record is outside the recursively cleaned
            # run. It is evidence for manual reconciliation, never deletion authority.
            with recovery.open("x", encoding="utf-8") as stream:
                stream.write(
                    json.dumps(
                        {
                            "schema": "docwen.managed-run-recovery.v1",
                            "root": str(self.root),
                            "rootIdentity": {"device": self.root_identity.st_dev, "inode": self.root_identity.st_ino},
                            "lease": self.lease,
                            "knownOutcome": outcome,
                            "executionError": (
                                {"type": type(execution_error).__name__, "message": str(execution_error)}
                                if execution_error is not None
                                else None
                            ),
                            "recordingError": {"type": type(recording_error).__name__, "message": str(recording_error)},
                            "phase": phase,
                            "error": str(primary),
                        }
                    )
                )
            primary.add_note(f"Managed run recovery record: {recovery}")
            print(f"Managed run recovery record: {recovery}", file=sys.stderr)
        except OSError as secondary:
            self.report_secondary(primary, secondary)


@contextmanager
def managed_run(parent: Path, *, prefix: str, owner: str, kind: str) -> Iterator[ManagedRun]:
    """Own allocation through teardown, preserving a primary failure on closeout errors."""
    parent = parent.resolve(strict=True)
    root = Path(tempfile.mkdtemp(prefix=prefix, dir=parent))
    try:
        root_identity = root.lstat()
    except OSError as primary:
        note = f"Cannot establish managed run identity; inspect {root} before cleanup"
        primary.add_note(note)
        print(note, file=sys.stderr)
        raise
    marker = root / ".docwen-temp-lease.json"
    marker_identity: os.stat_result | None = None

    def check_root() -> None:
        if root.parent != parent or root.resolve(strict=True) != root or not _same_entry(root, root_identity):
            raise ValueError(f"managed_run_identity_changed:{root}")

    def report_secondary(primary: BaseException, secondary: BaseException) -> None:
        note = f"Managed run closeout failed; inspect {root}: {secondary}"
        primary.add_note(note)
        print(note, file=sys.stderr)

    try:
        lease = lease_payload(root, owner=owner, kind=kind)
        with marker.open("x", encoding="utf-8") as stream:
            marker_identity = os.fstat(stream.fileno())
            json.dump(lease, stream)
    except BaseException as primary:
        try:
            check_root()
            # Only remove our own partial marker and an otherwise empty directory.
            # Never recursively erase an incompletely initialized run.
            if marker_identity is not None and _same_entry(marker, marker_identity):
                marker.unlink()
            root.rmdir()
        except (OSError, ValueError) as secondary:
            report_secondary(primary, secondary)
        raise

    assert marker_identity is not None
    captured = OwnedRunLease(root, root_identity, marker_identity, lease, owner, json.dumps(lease).encode("utf-8"))
    save = captured.save
    record_recovery = captured.record_recovery

    run = ManagedRun(root)
    try:
        yield run
    except BaseException as primary:
        outcome = "retained-interrupted" if isinstance(primary, KeyboardInterrupt) else "retained-failure"
        try:
            save(outcome)
        except (OSError, ValueError) as secondary:
            report_secondary(primary, secondary)
            record_recovery(
                primary, "failure-recording", outcome=outcome, recording_error=secondary, execution_error=primary
            )
        raise
    else:
        try:
            save(run.state)
        except (OSError, ValueError) as primary:
            record_recovery(primary, "terminal-recording", outcome=run.state, recording_error=primary)
            raise
        if run.state == "completed-success":
            root_pinned = False
            cleanup_failure_recorded = False
            try:
                with captured.pin_root():
                    root_pinned = True
                    try:
                        check_root()
                        shutil.rmtree(root)
                    except (OSError, ValueError) as primary:
                        try:
                            save("retained-cleanup-failure", restore_missing=True)
                        except (OSError, ValueError) as secondary:
                            report_secondary(primary, secondary)
                            record_recovery(primary, "cleanup-recording", outcome=run.state, recording_error=secondary)
                        cleanup_failure_recorded = True
                        raise
            except (OSError, ValueError) as primary:
                if not cleanup_failure_recorded:
                    # No trusted root pin: do not write a possibly foreign marker.
                    # The external hold protects even a previous success lease.
                    record_recovery(
                        primary,
                        "cleanup-close" if root_pinned else "cleanup-identity",
                        outcome=run.state,
                        recording_error=primary,
                    )
                raise


def lease_payload(root: Path, *, owner: str, kind: str, state: str = "active") -> dict[str, Any]:
    if not owner.startswith("docwen.") or not kind or state not in STATES:
        raise ValueError("invalid_run_lease_identity_or_state")
    identity = observe(os.getpid()).identity
    if os.name == "nt" and identity is None:
        raise OSError("cannot_identify_lease_owner_process")
    return {
        "schemaVersion": 1,
        "owner": owner,
        "kind": kind,
        "pid": os.getpid(),
        **({"processIdentity": identity} if identity is not None else {}),
        "createdAt": datetime.now(UTC).isoformat().replace("+00:00", "Z"),
        "state": state,
        "root": str(root.resolve(strict=True)),
    }


def transition(payload: dict[str, Any], *, root: Path, owner: str, state: str) -> None:
    if state not in STATES or payload.get("owner") != owner or payload.get("root") != str(root.resolve(strict=True)):
        raise ValueError("run_lease_transition_identity_or_state_mismatch")
    previous = payload.get("state")
    now = datetime.now(UTC)
    payload["state"] = state
    payload["updatedAt"] = now.isoformat().replace("+00:00", "Z")
    if state != "active":
        payload.setdefault("finishedAt", payload["updatedAt"])
        # Cleanup failure must not erase the run's actual success/failure outcome.
        if state in {"completed-success", "retained-failure", "retained-interrupted"}:
            payload.setdefault("outcome", state)
        if state == "retained-cleanup-failure" and previous in {"retained-failure", "retained-interrupted"}:
            payload.setdefault("outcome", previous)
    if state == "retained-manual":
        payload.setdefault(
            "retention",
            {
                "owner": owner,
                "reason": "explicit_keep_runtime_request",
                "reviewAfter": (now + timedelta(hours=72)).isoformat().replace("+00:00", "Z"),
            },
        )


def manual_retention_observation(payload: dict[str, Any], *, now: datetime) -> str:
    """Expiry requests review, never automatic deletion of an intentional hold."""
    retention = payload.get("retention")
    if not isinstance(retention, dict) or not retention.get("owner") or not retention.get("reason"):
        return "manual_retention_metadata_missing"
    try:
        deadline = datetime.fromisoformat(str(retention["reviewAfter"]).replace("Z", "+00:00"))
        if deadline.tzinfo is None:
            raise ValueError("timezone required")
    except (KeyError, ValueError):
        return "manual_retention_metadata_missing"
    return "manual_retention_review_due" if now >= deadline else "manual_retention_held"
