"""Disposable Shell doubles; their candidate paths stay inside pytest ownership."""

from __future__ import annotations

import os
import struct
from collections.abc import Callable
from pathlib import Path

import pytest
from tools import recycle

from tests.support.subprocess_runner import run_subprocess


def directory_link(link: Path, target: Path) -> None:
    try:
        link.symlink_to(target, target_is_directory=True)
    except OSError as error:
        if os.name != "nt":
            pytest.skip(f"directory symlink unavailable: {error}")
        result = run_subprocess(
            ["cmd.exe", "/d", "/c", "mklink", "/J", str(link), str(target)], capture_output=True, text=True
        )
        if result.returncode:
            pytest.skip(f"directory junction unavailable: {result.stderr}")


def fake_bin(tmp_path: Path, monkeypatch: pytest.MonkeyPatch, *, after_move: Callable[[Path], None] | None = None):
    root = tmp_path / "$Recycle.Bin"
    owner = root / "test-owner"
    owner.mkdir(parents=True)
    monkeypatch.setattr(recycle, "_recycle_root", lambda path: root)
    calls = []

    def perform(target: Path) -> None:
        suffix = f"fixture{len(calls)}"
        metadata = owner / ("$I" + suffix)
        payload = owner / ("$R" + suffix)
        original = str(target) + "\0"
        metadata.write_bytes(struct.pack("<QQQI", 2, 0, 0, len(original)) + original.encode("utf-16-le"))
        target.rename(payload)
        calls.append({"metadata": str(metadata), "payload": str(payload)})
        if after_move is not None:
            after_move(payload)

    monkeypatch.setattr(recycle, "_perform_recycle", perform)
    return owner, calls
