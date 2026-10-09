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

    def same_entry(path: Path, expected: os.stat_result) -> bool:
        actual = path.lstat()
        return (
            (actual.st_dev, actual.st_ino) == (expected.st_dev, expected.st_ino)
            and not stat.S_ISLNK(actual.st_mode)
            and not (getattr(actual, "st_file_attributes", 0) & getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0))
        )

    def check_root() -> None:
        if root.parent != parent or root.resolve(strict=True) != root or not same_entry(root, root_identity):
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
            if marker_identity is not None and same_entry(marker, marker_identity):
                marker.unlink()
            root.rmdir()
        except (OSError, ValueError) as secondary:
            report_secondary(primary, secondary)
        raise

    def save(state: str, *, restore_missing: bool = False) -> None:
        nonlocal marker_identity
        check_root()
        try:
            marker_matches = marker_identity is not None and same_entry(marker, marker_identity)
        except FileNotFoundError:
            if not restore_missing:
                raise
            # Partial recursive cleanup may have removed our marker. The root
            # still has its original identity; exclusive creation cannot replace
            # a new owner's marker that appears after this check.
            transition(lease, root=root, owner=owner, state=state)
            with marker.open("x", encoding="utf-8") as stream:
                marker_identity = os.fstat(stream.fileno())
                stream.write(json.dumps(lease))
                stream.flush()
                os.fsync(stream.fileno())
            return
        if not marker_matches:
            raise ValueError(f"managed_run_marker_changed:{marker}")
        transition(lease, root=root, owner=owner, state=state)
        temporary = root / ".docwen-temp-lease.next"
        with temporary.open("x", encoding="utf-8") as stream:
            next_identity = os.fstat(stream.fileno())
            json.dump(lease, stream)
            stream.flush()
            os.fsync(stream.fileno())
        temporary.replace(marker)
        marker_identity = next_identity

    def record_recovery(
        primary: BaseException,
        phase: str,
        *,
        outcome: str,
        recording_error: BaseException,
        execution_error: BaseException | None = None,
    ) -> None:
        recovery = parent / f"{root.name}.recovery.json"
        try:
            # This compact recovery record is outside the recursively cleaned
            # run. It is evidence for manual reconciliation, never deletion authority.
            with recovery.open("x", encoding="utf-8") as stream:
                stream.write(
                    json.dumps(
                        {
                            "schema": "docwen.managed-run-recovery.v1",
                            "root": str(root),
                            "rootIdentity": {"device": root_identity.st_dev, "inode": root_identity.st_ino},
                            "lease": lease,
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
            report_secondary(primary, secondary)

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
            try:
                check_root()
                shutil.rmtree(root)
            except (OSError, ValueError) as primary:
                try:
                    save("retained-cleanup-failure", restore_missing=True)
                except (OSError, ValueError) as secondary:
                    report_secondary(primary, secondary)
                    record_recovery(primary, "cleanup-recording", outcome=run.state, recording_error=secondary)
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
