from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pytest
from tools.export_consumer_contracts import export_snapshot, snapshot_files

pytestmark = pytest.mark.contract


def test_snapshot_normalizes_only_json_and_binds_original_blobs(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    commit = "a" * 40
    manifest = {"schemas": [{"path": "schemas/example.json"}], "fixtures": []}
    blobs = {
        "contracts/conformance-manifest.json": json.dumps(manifest).encode(),
        "contracts/schemas/example.json": '{"title":"原合同", "required":[]}\n'.encode(),
        "LICENSE": b"Original license\n",
    }

    def git_blob(_repo: Path, *args: str) -> bytes:
        if args[0] == "rev-parse":
            return commit.encode()
        assert args[0] == "show"
        revision, name = args[1].split(":", 1)
        assert revision == commit
        return blobs[name]

    monkeypatch.setattr("tools.export_consumer_contracts._git", git_blob)
    files = snapshot_files(tmp_path)
    metadata = json.loads(files["snapshot.json"])
    assert metadata["source_commit"] == commit
    for record in metadata["files"]:
        assert record["source_sha256"] == hashlib.sha256(blobs[record["source_path"]]).hexdigest()
        assert record["sha256"] == hashlib.sha256(files[record["path"]]).hexdigest()
    assert files["LICENSE"] == blobs["LICENSE"]
    assert files["schemas/example.json"] != blobs["contracts/schemas/example.json"]
    assert json.loads(files["schemas/example.json"]) == json.loads(blobs["contracts/schemas/example.json"])
    destination = tmp_path / "snapshot"
    export_snapshot(destination, files)
    export_snapshot(destination, files, check=True)
    with pytest.raises(FileExistsError):
        export_snapshot(destination, files)


@pytest.mark.parametrize("change", ["modify", "missing", "extra"])
def test_snapshot_check_rejects_drift_without_repairing_it(tmp_path: Path, change: str) -> None:
    destination = tmp_path / "snapshot"
    files = {"schemas/a.json": b"{}\n"}
    export_snapshot(destination, files)
    target = destination / "schemas/a.json"
    if change == "modify":
        target.write_bytes(b"[]\n")
    elif change == "missing":
        target.unlink()
    else:
        (destination / "extra.json").write_bytes(b"{}\n")
    before = {path.relative_to(destination): path.read_bytes() for path in destination.rglob("*") if path.is_file()}
    with pytest.raises(ValueError, match="Snapshot"):
        export_snapshot(destination, files, check=True)
    assert before == {
        path.relative_to(destination): path.read_bytes() for path in destination.rglob("*") if path.is_file()
    }


@pytest.mark.parametrize("path", ["../outside.json", "/outside.json", "a\\b.json", "C:/outside.json"])
def test_snapshot_rejects_manifest_path_escape(monkeypatch: pytest.MonkeyPatch, tmp_path: Path, path: str) -> None:
    manifest = json.dumps({"schemas": [{"path": path}], "fixtures": []}).encode()
    monkeypatch.setattr(
        "tools.export_consumer_contracts._git", lambda _repo, *args: b"a" * 40 if args[0] == "rev-parse" else manifest
    )
    with pytest.raises(ValueError, match="Unsafe contract path"):
        snapshot_files(tmp_path)
    assert not list(tmp_path.iterdir())
