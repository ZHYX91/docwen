"""Canonical identities and states for engineering run leases."""

from __future__ import annotations

import os
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
