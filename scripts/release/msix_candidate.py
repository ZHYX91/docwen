"""Extract a verified portable candidate for the independent MSIX channel."""

from __future__ import annotations

import re
import shutil
import stat
import zipfile
from pathlib import Path, PurePosixPath
from typing import Any

from scripts.release.publication_contract import PublicationError, file_identity, read_object, require


def extract_candidate(archive_path: Path, receipt_path: Path, output: Path, *, version: str) -> dict[str, Any]:
    receipt = read_object(receipt_path)
    require(
        receipt.get("stage") == "candidate-verified" and receipt.get("provenance") == "verified",
        "MSIX requires an inspected candidate",
    )
    require(receipt.get("version") == version, "MSIX candidate version mismatch")
    expected = receipt.get("assets", {}).get("DocWen-windows-x64.zip")
    require(file_identity(archive_path) == expected, "MSIX portable archive identity mismatch")
    require(not output.exists(), "MSIX extraction output must be new")
    with zipfile.ZipFile(archive_path) as archive:
        entries = archive.infolist()
        require(
            0 < len(entries) <= 100000 and sum(entry.file_size for entry in entries) <= 4 * 1024**3,
            "MSIX portable inventory exceeds limits",
        )
        seen: set[str] = set()
        for entry in entries:
            name = entry.filename
            require(entry.orig_filename == name, "MSIX portable path was normalized")
            path = PurePosixPath(name)
            parts = path.parts
            require(bool(parts) and str(path) == name and not path.is_absolute(), "MSIX portable path invalid")
            require(
                all(
                    part not in {".", ".."}
                    and not part.endswith((".", " "))
                    and not re.search(r'[\\:<>"|?*\x00-\x1f]', part)
                    for part in parts
                ),
                "MSIX portable path invalid",
            )
            require(
                all(re.fullmatch(r"(?i)(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\..*)?", part) is None for part in parts),
                "MSIX portable device path rejected",
            )
            require(
                not entry.is_dir() and stat.S_ISREG(entry.external_attr >> 16), "MSIX portable member must be regular"
            )
            key = name.casefold()
            require(key not in seen, "MSIX portable duplicate path")
            seen.add(key)
        require({"docwen.exe", "docwencli.exe"}.issubset(seen), "MSIX portable executables missing")
        output.mkdir()
        for entry in entries:
            target = output.joinpath(*PurePosixPath(entry.filename).parts)
            target.parent.mkdir(parents=True, exist_ok=True)
            with archive.open(entry) as source, target.open("xb") as destination:
                shutil.copyfileobj(source, destination)
    require(file_identity(archive_path) == expected, "MSIX portable archive changed during extraction")
    try:
        return {
            name: receipt[name]
            for name in (
                "repository",
                "version",
                "sourceCommit",
                "manifestSha256",
                "artifactId",
                "artifactDigest",
                "origin",
            )
        } | {"portableArchive": expected}
    except KeyError as error:
        raise PublicationError("MSIX candidate receipt identity incomplete") from error
