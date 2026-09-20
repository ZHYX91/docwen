"""Export a self-contained, revision-bound copy of the normative wire test data."""

from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
from pathlib import Path, PurePosixPath
from typing import Any

REPO_ROOT = Path(__file__).resolve().parents[1]


def _git(repo: Path, *args: str) -> bytes:
    return subprocess.check_output(["git", "-C", str(repo), *args])


def _json_bytes(value: Any) -> bytes:
    return (json.dumps(value, ensure_ascii=False, indent=2) + "\n").encode("utf-8")


def _relative_path(value: str) -> str:
    path = PurePosixPath(value)
    if "\\" in value or ":" in value or path.is_absolute() or any(part in {"", ".", ".."} for part in value.split("/")):
        raise ValueError(f"Unsafe contract path: {value}")
    return value


def snapshot_files(repo: Path, revision: str = "HEAD") -> dict[str, bytes]:
    """Read committed blobs only; normalize JSON whitespace without changing its data."""
    commit = _git(repo, "rev-parse", "--verify", f"{revision}^{{commit}}").decode().strip()
    manifest_path = "contracts/conformance-manifest.json"
    manifest_bytes = _git(repo, "show", f"{commit}:{manifest_path}")
    manifest = json.loads(manifest_bytes)
    source_paths = [manifest_path, "LICENSE"]
    source_paths.extend(
        f"contracts/{_relative_path(item['path'])}" for item in [*manifest["schemas"], *manifest["fixtures"]]
    )
    if len(set(source_paths)) != len(source_paths):
        raise ValueError("Duplicate contract inventory path")
    exported: dict[str, bytes] = {}
    records = []
    for source_path in sorted(source_paths):
        original = _git(repo, "show", f"{commit}:{source_path}")
        relative = source_path.removeprefix("contracts/")
        normalized = _json_bytes(json.loads(original)) if relative.endswith(".json") else original
        exported[relative] = normalized
        records.append(
            {
                "path": relative,
                "source_path": source_path,
                "source_sha256": hashlib.sha256(original).hexdigest(),
                "sha256": hashlib.sha256(normalized).hexdigest(),
                "size_bytes": len(normalized),
            }
        )
    exported["snapshot.json"] = _json_bytes(
        {
            "schema": "docwen.consumer_contract_snapshot.v1",
            "source_repository": "https://github.com/ZHYX91/DocWen",
            "source_commit": commit,
            "normalization": "JSON: UTF-8, two-space indentation, one LF; data unchanged. Other files: original Git blob bytes.",
            "license": "AGPL-3.0-or-later (vendored test data only; see LICENSE)",
            "files": records,
        }
    )
    return exported


def export_snapshot(destination: Path, files: dict[str, bytes], *, check: bool = False) -> None:
    """Create a new snapshot or compare a complete existing snapshot without modifying it."""
    destination = destination.absolute()
    if any(parent.is_symlink() for parent in (destination, *destination.parents)):
        raise ValueError("Snapshot destination must not contain symbolic links")
    if check:
        actual = {path.relative_to(destination).as_posix() for path in destination.rglob("*") if path.is_file()}
        if actual != set(files):
            raise ValueError("Snapshot file inventory differs from the selected source revision")
        for relative, expected in files.items():
            path = destination / relative
            if path.is_symlink() or path.read_bytes() != expected:
                raise ValueError(f"Snapshot differs from the selected source revision: {relative}")
        return
    destination.mkdir(parents=True, exist_ok=False)
    for relative, content in files.items():
        path = destination / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("xb") as stream:
            stream.write(content)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("destination", type=Path)
    parser.add_argument(
        "--revision", default="HEAD", help="Committed source revision; working-tree edits are never exported"
    )
    parser.add_argument("--check", action="store_true", help="Verify an existing snapshot without writing")
    args = parser.parse_args()
    files = snapshot_files(REPO_ROOT, args.revision)
    export_snapshot(args.destination, files, check=args.check)
    print(f"{'Verified' if args.check else 'Exported'} {len(files)} contract snapshot files")


if __name__ == "__main__":
    main()
