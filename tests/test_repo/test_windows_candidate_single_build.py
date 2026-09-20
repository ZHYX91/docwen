from __future__ import annotations

import hashlib
import json
import sys
import zipfile
from collections.abc import Sequence
from pathlib import Path

import pytest
from scripts.release import build_production_candidate as production

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]


@pytest.mark.parametrize("gui_bytes", [b"first source version", b"changed source version"])
def test_candidate_records_one_build_without_precomputed_output_hashes(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, gui_bytes: bytes
) -> None:
    repo = tmp_path / "repo"
    manifest_path = repo / "release" / "windows-production-manifest.v1.json"
    manifest_path.parent.mkdir(parents=True)
    manifest_path.write_bytes((ROOT / "release/windows-production-manifest.v1.json").read_bytes())
    manifest = production.read_manifest(manifest_path)
    assert "allowlist" not in manifest["payload"]
    work, output = tmp_path / "work", tmp_path / "output"
    python = Path(sys.executable).resolve()
    uv = tmp_path / "uv"
    uv.write_bytes(b"verified-toolchain-placeholder")
    build_calls: list[Sequence[str]] = []

    def fake_git(_repo: Path, *args: str) -> str:
        if args[0] == "status":
            return ""
        if args[0] == "show":
            return "1704067200"
        return "b" * 40 if args[-1] == "HEAD^{tree}" else "a" * 40

    def fake_run(command: Sequence[str], *, cwd: Path, env: dict[str, str] | None = None) -> str:
        if command[:2] == ["git", "clone"]:
            (work / "source").mkdir()
        elif command[:2] == ["git", "checkout"]:
            pass
        else:
            assert command[1] == "scripts/build/build.py"
            assert env is not None and env["UV_OFFLINE"] == "1"
            build_calls.append(command)
            deploy = cwd / "dist" / f"DocWen_v{manifest['product']['version']}_win-x64"
            deploy.mkdir(parents=True)
            (deploy / "DocWen.exe").write_bytes(gui_bytes)
            (deploy / "DocWenCLI.exe").write_bytes(b"cli")
        return "primitive build completed"

    monkeypatch.setattr(production, "git", fake_git)
    monkeypatch.setattr(production, "run", fake_run)
    monkeypatch.setattr(production, "verify_source_contracts", lambda *_: {"contract": "verified"})
    monkeypatch.setattr(production, "_verify_toolchain", lambda *_: ({"toolchain": "verified"}, python))
    monkeypatch.setattr(production, "sync_project_environment", lambda *_: python)
    monkeypatch.setattr(production, "windows_build_environment", lambda env, *_: env)
    monkeypatch.setattr(production, "normalize_packaged_msvc_runtime", lambda *_: {})
    args = production.parser().parse_args(
        [
            "--repo",
            str(repo),
            "--manifest",
            str(manifest_path),
            "--python",
            str(python),
            "--uv",
            str(uv),
            "--work-root",
            str(work),
            "--output-root",
            str(output),
        ]
    )

    result = production.build(args)

    assert len(build_calls) == 1
    payload_manifest_path = output / "evidence" / "payload-manifest.json"
    payload_manifest = json.loads(payload_manifest_path.read_bytes())
    assert payload_manifest["files"][0] == {
        "path": "DocWen.exe",
        "bytes": len(gui_bytes),
        "executable": True,
        "sha256": hashlib.sha256(gui_bytes).hexdigest(),
    }
    assert result["payloadManifestSha256"] == production.sha256_file(payload_manifest_path)
    archive = output / result["offlineZip"]["name"]
    assert result["offlineZip"]["sha256"] == production.sha256_file(archive)
    with zipfile.ZipFile(archive) as packaged:
        assert packaged.namelist() == ["DocWen.exe", "DocWenCLI.exe"]
        assert packaged.read("DocWen.exe") == gui_bytes
    assert not work.exists()
    assert not (output / production.PRODUCTION_OUTPUT_LEASE).exists()
