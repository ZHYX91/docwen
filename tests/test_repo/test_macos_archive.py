"""Protect archive containment, runtime modes and post-extract integrity."""

from __future__ import annotations

import os
import tarfile
from pathlib import Path

import pytest
from scripts.release import macos_package
from scripts.release.macos_archive import archive_payload, inventory, verify_extracted, write_inventory
from scripts.release.macos_package import read_contract
from scripts.release.publication_contract import PublicationError

pytestmark = pytest.mark.contract


def payload(tmp_path: Path) -> Path:
    root = tmp_path / "payload"
    root.mkdir()
    (root / "DocWenCLI").write_bytes(b"executable")
    (root / "DocWenCLI").chmod(0o755)
    (root / "文档.txt").write_text("Unicode content", encoding="utf-8")
    return root


def test_archive_roundtrip_preserves_content_and_executable_modes(tmp_path: Path) -> None:
    root = payload(tmp_path)
    write_inventory(root)
    archive = tmp_path / "app.tar.gz"
    archive_payload(root, archive, top="DocWen-test", epoch=1704067200)
    with tarfile.open(archive) as bundle:
        entry = bundle.getmember("DocWen-test/DocWenCLI")
        if os.name != "nt":
            assert entry.mode == 0o755
        bundle.extractall(tmp_path / "unpacked", filter="data")
    extracted = tmp_path / "unpacked/DocWen-test"
    verify_extracted(extracted)
    (extracted / "文档.txt").write_bytes(b"tampered")
    with pytest.raises(PublicationError, match="identity mismatch"):
        verify_extracted(extracted)


@pytest.mark.skipif(os.name == "nt", reason="POSIX symlink contract runs on both release POSIX platforms")
@pytest.mark.parametrize("target", ["/etc/passwd", "../outside", "missing", "loop", "."])
def test_reject_unsafe_links_before_archiving(tmp_path: Path, target: str) -> None:
    root = payload(tmp_path)
    (tmp_path / "outside").write_bytes(b"outside")
    (root / "loop").symlink_to(target)
    with pytest.raises((PublicationError, ValueError)):
        inventory(root)


@pytest.mark.skipif(os.name == "nt", reason="POSIX symlink preservation")
def test_internal_framework_directory_links_are_preserved(tmp_path: Path) -> None:
    root = payload(tmp_path)
    (root / "Versions/A").mkdir(parents=True)
    (root / "Versions/A/library").write_bytes(b"library")
    (root / "Versions/Current").symlink_to("A", target_is_directory=True)
    before = inventory(root)
    archive_payload(root, tmp_path / "app.tar.gz", top="App", epoch=1)
    with tarfile.open(tmp_path / "app.tar.gz") as bundle:
        assert bundle.getmember("App/Versions/Current").issym()
        bundle.extractall(tmp_path / "unpacked", filter="data")
    assert inventory(tmp_path / "unpacked/App") == before


def test_macos_production_contract_has_only_ad_hoc_arm64_assets() -> None:
    contract = read_contract(Path("release/macos-production-manifest.v1.json"))
    assert contract["signing"] == {"identity": "-", "developerId": False, "notarized": False}
    assert len(contract["artifacts"]) == 2


def test_contract_cannot_silently_enable_identity_signing(tmp_path: Path) -> None:
    import json

    contract = json.loads(Path("release/macos-production-manifest.v1.json").read_text(encoding="utf-8"))
    contract["signing"]["identity"] = "Developer ID Application: unexpected"
    path = tmp_path / "contract.json"
    path.write_text(json.dumps(contract), encoding="utf-8")
    with pytest.raises(PublicationError, match="ad-hoc"):
        read_contract(path)


@pytest.mark.parametrize("sealed", [False, True])
def test_framework_verification_requires_the_complete_bundle_resource_seal(tmp_path, monkeypatch, sealed):
    framework = tmp_path / "QtCore.framework"
    binary = framework / "Versions/A/QtCore"
    binary.parent.mkdir(parents=True)
    binary.write_bytes(b"\xcf\xfa\xed\xfecode")
    calls = []

    def run(arguments):
        calls.append(arguments)
        if "lipo" in arguments[0]:
            return "arm64"
        assert arguments[-1] == str(framework)
        assert "--deep" in arguments and "--strict" in arguments
        if not sealed:
            raise PublicationError("missing framework resource seal")
        return ""

    monkeypatch.setattr(macos_package, "command", run)
    if sealed:
        assert macos_package.verify_signatures(tmp_path)["machOCount"] == 1
    else:
        with pytest.raises(PublicationError, match="resource seal"):
            macos_package.verify_signatures(tmp_path)
    assert len(calls) == 2
