"""Archive immutable macOS payloads, preserving executable modes and internal links."""

from __future__ import annotations

import gzip
import hashlib
import json
import os
import stat
import tarfile
from pathlib import Path, PurePosixPath
from typing import Any

from scripts.release.publication_contract import canonical_json, file_identity, require


def inventory(root: Path) -> dict[str, dict[str, Any]]:
    """Never follow directory links; reject escaping, dangling and cyclic links."""
    require(root.is_dir() and not root.is_symlink(), "plain payload directory required")
    resolved_root = root.resolve(strict=True)
    records: dict[str, dict[str, Any]] = {}
    for current, directories, files in os.walk(root, followlinks=False):
        for name in sorted(directories + files):
            path = Path(current) / name
            relative = path.relative_to(root).as_posix()
            require(
                "\\" not in relative and all(ord(c) >= 32 and ord(c) != 127 for c in relative), "unsafe payload path"
            )
            mode = path.lstat().st_mode
            if stat.S_ISLNK(mode):
                target = os.readlink(path)
                require(not PurePosixPath(target).is_absolute() and "\\" not in target, "absolute link rejected")
                try:
                    destination = path.resolve(strict=True)
                except (OSError, RuntimeError) as error:
                    raise ValueError(f"invalid payload link: {relative}") from error
                require(destination.is_relative_to(resolved_root), "escaping payload link")
                require(destination.is_file() or destination.is_dir(), "special link target")
                # A link to an ancestor is a traversal cycle even if resolve() terminates.
                require(not path.parent.resolve().is_relative_to(destination), "cyclic directory link")
                records[relative] = {"type": "symlink", "target": target, "mode": 0o777}
            elif stat.S_ISDIR(mode):
                records[relative] = {"type": "directory", "mode": 0o755}
            else:
                require(stat.S_ISREG(mode), "special payload file rejected")
                records[relative] = {"type": "file", "mode": 0o755 if mode & 0o111 else 0o644, **file_identity(path)}
    return dict(sorted(records.items(), key=lambda item: item[0].encode("utf-8")))


def archive_payload(root: Path, destination: Path, *, top: str, epoch: int) -> dict[str, Any]:
    require(not destination.exists(), "archive destination already exists")
    require(top and PurePosixPath(top).name == top and top not in {".", ".."}, "invalid archive root")
    before = inventory(root)
    with (
        destination.open("xb") as raw,
        gzip.GzipFile(filename="", mode="wb", fileobj=raw, mtime=epoch) as compressed,
        tarfile.open(fileobj=compressed, mode="w", format=tarfile.PAX_FORMAT) as archive,
    ):
        for relative, record in [("", {"type": "directory", "mode": 0o755}), *before.items()]:
            info = tarfile.TarInfo(top + ("/" + relative if relative else ""))
            info.mtime, info.mode = epoch, record["mode"]
            if record["type"] == "directory":
                info.type = tarfile.DIRTYPE
                archive.addfile(info)
            elif record["type"] == "symlink":
                info.type, info.linkname = tarfile.SYMTYPE, record["target"]
                archive.addfile(info)
            else:
                info.size = record["bytes"]
                with (root / relative).open("rb") as stream:
                    archive.addfile(info, stream)
    require(inventory(root) == before, "payload changed while archiving")
    verify_archive(destination, top=top, expected=before)
    return file_identity(destination)


def verify_archive(archive_path: Path, *, top: str, expected: dict[str, dict[str, Any]]) -> None:
    """Verify archive bytes and headers before extraction; no filesystem writes."""
    seen: set[str] = set()
    with tarfile.open(archive_path, "r:gz") as archive:
        for member in archive:
            require(member.name not in seen, "duplicate archive member")
            seen.add(member.name)
            path = PurePosixPath(member.name)
            require(path.parts[0] == top and not path.is_absolute() and ".." not in path.parts, "unsafe archive member")
            require(member.uid == member.gid == 0 and not member.uname and not member.gname, "archive ownership drift")
            if member.name == top:
                require(member.isdir(), "archive root is not a directory")
                continue
            relative = path.relative_to(top).as_posix()
            require(relative in expected, "unexpected archive member")
            record = expected[relative]
            require(member.mode == record["mode"], "archive mode mismatch")
            if record["type"] == "directory":
                require(member.isdir(), "archive directory mismatch")
            elif record["type"] == "symlink":
                require(member.issym() and member.linkname == record["target"], "archive link mismatch")
            else:
                require(member.isfile() and member.size == record["bytes"], "archive file mismatch")
                stream = archive.extractfile(member)
                require(stream is not None, "missing archive stream")
                assert stream is not None
                with stream:
                    require(
                        hashlib.file_digest(stream, "sha256").hexdigest() == record["sha256"], "archive hash mismatch"
                    )
    require(seen == {top} | {top + "/" + name for name in expected}, "incomplete archive")


def write_inventory(root: Path) -> None:
    records = inventory(root)
    (root / "manifest.json").write_bytes(canonical_json({"schema": "docwen.macos.payload.v1", "entries": records}))
    checksums = {**records, "manifest.json": {"type": "file", **file_identity(root / "manifest.json")}}
    (root / "SHA256SUMS.txt").write_text(
        "".join(f"{item['sha256']}  {name}\n" for name, item in sorted(checksums.items()) if item["type"] == "file"),
        encoding="utf-8",
        newline="\n",
    )


def verify_extracted(root: Path) -> None:
    records = json.loads((root / "manifest.json").read_text(encoding="utf-8"))
    require(records["schema"] == "docwen.macos.payload.v1", "invalid payload manifest")
    actual = inventory(root)
    actual.pop("manifest.json")
    actual.pop("SHA256SUMS.txt")
    require(actual == records["entries"], "extracted payload identity mismatch")
    checksums = {**actual, "manifest.json": {"type": "file", **file_identity(root / "manifest.json")}}
    expected = "".join(
        f"{item['sha256']}  {name}\n" for name, item in sorted(checksums.items()) if item["type"] == "file"
    )
    require((root / "SHA256SUMS.txt").read_text(encoding="utf-8") == expected, "payload checksum inventory mismatch")
