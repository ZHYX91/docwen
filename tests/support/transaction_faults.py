"""Narrow real-filesystem fault injection for committed journal cleanup.

The replace and unlink operations run normally. Tests inspect the final bytes,
identity and catalog independently; this helper only places an I/O failure at
one already committed cleanup boundary.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path

import pytest


@dataclass
class CommittedCleanupFault:
    committed_envelope: bytes | None = None
    removed: bool = False
    injected: bool = False


def fail_committed_journal_cleanup_once(
    monkeypatch: pytest.MonkeyPatch,
    journal: Path,
    *,
    boundary: str = "directory_sync",
) -> CommittedCleanupFault:
    """Fail either unlink or its following barrier after actual COMMITTED replace."""
    from docwen_runtime import toml_io

    if boundary not in {"unlink", "directory_sync"}:
        raise ValueError(f"Unknown cleanup fault boundary: {boundary}")
    evidence = CommittedCleanupFault()
    original_replace = toml_io.os.replace
    original_unlink = Path.unlink
    original_sync = toml_io._sync_directory

    def replace(source, destination) -> None:
        original_replace(source, destination)
        if Path(destination) == journal:
            payload = journal.read_bytes()
            if json.loads(payload)["payload"]["state"] == "COMMITTED":
                evidence.committed_envelope = payload

    def unlink(path: Path, *args, **kwargs) -> None:
        if (
            path == journal
            and evidence.committed_envelope is not None
            and boundary == "unlink"
            and not evidence.injected
        ):
            evidence.injected = True
            raise OSError("injected committed journal unlink failure")
        original_unlink(path, *args, **kwargs)
        if path == journal and evidence.committed_envelope is not None:
            evidence.removed = True

    def sync(path) -> None:
        if Path(path) == journal.parent and boundary == "directory_sync" and evidence.removed and not evidence.injected:
            assert not journal.exists()
            evidence.injected = True
            raise OSError("injected committed journal post-unlink directory barrier failure")
        original_sync(path)

    monkeypatch.setattr(toml_io.os, "replace", replace)
    monkeypatch.setattr(Path, "unlink", unlink)
    monkeypatch.setattr(toml_io, "_sync_directory", sync)
    return evidence
