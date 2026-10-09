"""Read back SDK output using decoded OPC names and exact staged bytes."""

from __future__ import annotations

import hashlib
import zipfile
from pathlib import Path, PurePosixPath
from urllib.parse import unquote

from scripts.release.publication_contract import require


def verify_layout(package_path: Path, staging: Path) -> int:
    expected = {path.relative_to(staging).as_posix(): path for path in staging.rglob("*") if path.is_file()}
    generated = {"AppxBlockMap.xml", "[Content_Types].xml"}
    with zipfile.ZipFile(package_path) as archive:
        members = {}
        folded = set()
        for entry in archive.infolist():
            name = unquote(entry.filename, encoding="utf-8", errors="strict")
            logical = PurePosixPath(name)
            require(
                not entry.is_dir()
                and not logical.is_absolute()
                and ".." not in logical.parts
                and "\\" not in name
                and str(logical) == name,
                "MSIX member path invalid",
            )
            require(name.casefold() not in folded, "MSIX decoded member collision")
            folded.add(name.casefold())
            members[name] = entry
        require(set(members) == set(expected) | generated, "MSIX staged inventory mismatch")
        for name, path in expected.items():
            with path.open("rb") as source, archive.open(members[name]) as packaged:
                require(
                    hashlib.file_digest(source, "sha256").digest() == hashlib.file_digest(packaged, "sha256").digest(),
                    f"MSIX staged content mismatch: {name}",
                )
    return len(expected)
