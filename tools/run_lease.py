"""Canonical identities and states for engineering run leases."""

from __future__ import annotations

import os
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

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
    return {
        "schemaVersion": 1,
        "owner": owner,
        "kind": kind,
        "pid": os.getpid(),
        "createdAt": datetime.now(UTC).isoformat().replace("+00:00", "Z"),
        "state": state,
        "root": str(root.resolve(strict=True)),
    }


def transition(payload: dict[str, Any], *, root: Path, owner: str, state: str) -> None:
    if state not in STATES or payload.get("owner") != owner or payload.get("root") != str(root.resolve(strict=True)):
        raise ValueError("run_lease_transition_identity_or_state_mismatch")
    payload["state"] = state
    payload["updatedAt"] = datetime.now(UTC).isoformat().replace("+00:00", "Z")
