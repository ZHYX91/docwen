from __future__ import annotations

import struct
from pathlib import Path
from types import SimpleNamespace

import pytest
from tools import lx_symlink, workspace_cleanup

pytestmark = pytest.mark.unit


def _buffer(target: str, version: int = 2) -> bytes:
    encoded = target.encode("utf-8")
    return struct.pack("<IHHI", lx_symlink.LX_SYMLINK_TAG, len(encoded) + 4, 0, version) + encoded


def test_decode_real_format() -> None:
    assert lx_symlink.decode_target(_buffer("libQt6Core.so.6.8.3")) == "libQt6Core.so.6.8.3"


@pytest.mark.parametrize("target", ["", "/etc/passwd", "../outside", "a/../../b", "C:secret", "a\\b", "x\0y"])
def test_reject_ambiguous_guest_paths(target: str) -> None:
    with pytest.raises(ValueError):
        lx_symlink.decode_target(_buffer(target))


def test_reject_malformed_buffer() -> None:
    for data in (_buffer("target", 1), _buffer("target")[:-1], b"x"):
        with pytest.raises(ValueError):
            lx_symlink.decode_target(data)


def test_lx_resolution_requires_plain_existing_target(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    link = tmp_path / "link"
    target = tmp_path / "target"
    target.write_text("payload")
    original = Path.lstat
    monkeypatch.setattr(
        Path,
        "lstat",
        lambda self: SimpleNamespace(st_reparse_tag=lx_symlink.LX_SYMLINK_TAG) if self == link else original(self),
    )
    monkeypatch.setattr(lx_symlink, "read_target", lambda path: "target")
    assert workspace_cleanup._resolve_reparse_target(link) == target
    target.unlink()
    with pytest.raises(workspace_cleanup.HousekeepingError):
        workspace_cleanup._resolve_reparse_target(link)
