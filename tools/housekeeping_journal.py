"""Durable per-target progress; never retry an ambiguous filesystem mutation."""

from __future__ import annotations

import json
import os
from contextlib import contextmanager
from pathlib import Path
from typing import Any


@contextmanager
def exclusive_apply(path: Path):
    """Use an OS lock, released even when the applying process dies."""
    with path.open("a+b") as stream:
        stream.seek(0, os.SEEK_END)
        if stream.tell() == 0:
            stream.write(b"\0")
            stream.flush()
        stream.seek(0)
        if os.name == "nt":
            import msvcrt

            msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
        else:
            import fcntl

            fcntl.flock(stream.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        try:
            yield
        finally:
            stream.seek(0)
            if os.name == "nt":
                msvcrt.locking(stream.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(stream.fileno(), fcntl.LOCK_UN)


class ApplyJournal:
    def __init__(self, path: Path, fingerprint: str):
        self.path = path
        self.payload: dict[str, Any] = {
            "schema": "docwen.housekeeping-progress.v1",
            "planFingerprint": fingerprint,
            "completed": [],
            "pending": None,
        }
        if path.exists():
            payload = json.loads(path.read_text(encoding="utf-8"))
            if (
                not isinstance(payload, dict)
                or payload.get("schema") != self.payload["schema"]
                or payload.get("planFingerprint") != fingerprint
                or not isinstance(payload.get("completed"), list)
                or "pending" not in payload
                or any(not isinstance(item, dict) for item in payload["completed"])
                or (payload["pending"] is not None and not isinstance(payload["pending"], dict))
            ):
                raise ValueError("housekeeping_journal_identity_mismatch")
            self.payload = payload

    def save(self) -> None:
        temporary = self.path.with_suffix(self.path.suffix + ".tmp")
        # Exclusive creation refuses a leftover or linked temporary file.
        with temporary.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(self.payload, stream, ensure_ascii=False, indent=2)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())
        temporary.replace(self.path)

    def begin(self, path: Path) -> None:
        self.payload["pending"] = {"path": str(path), "state": "outcome-unconfirmed"}
        self.save()

    def complete(self, result: dict[str, Any]) -> None:
        self.payload["completed"].append(result)
        self.payload["pending"] = None
        self.save()

    def failed(self, error: BaseException) -> None:
        self.payload["pending"]["error"] = f"{type(error).__name__}:{error}"
        self.save()
