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
        if "otool" in arguments[0]:
            assert arguments == ["/usr/bin/otool", "-arch", "arm64", "-l", str(binary)]
            return "Load command 0\n      cmd LC_BUILD_VERSION\n platform MACOS\n    minos 14.0\n      sdk 15.0\n"
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
    assert len(calls) == (3 if sealed else 2)


@pytest.mark.parametrize(
    ("metadata", "expected"),
    [
        ("cmd LC_BUILD_VERSION\n platform MACOS\n minos 14.0\n sdk 15.0\n ntools 1\n tool LD\n version 1167.5", "14.0"),
        ("cmd LC_BUILD_VERSION\n platform 1\n minos 11.3\n sdk 15.0", "11.3"),
        ("cmd LC_BUILD_VERSION\n platform 1\n minos 14.0.0\n sdk 15.0", "14.0"),
        ("cmd LC_VERSION_MIN_MACOSX\n version 10.15.7\n sdk 15.0", "10.15.7"),
    ],
)
@pytest.mark.parametrize("magic", sorted(macos_package.MACHO_MAGICS))
def test_arm64_deployment_floor_uses_minimum_not_sdk(tmp_path, monkeypatch, metadata, expected, magic):
    binary = tmp_path / "nested/library.dylib"
    binary.parent.mkdir()
    binary.write_bytes(magic + b"code")
    original = binary.read_bytes()

    def run(arguments):
        if arguments[0] == "/usr/bin/lipo":
            return "x86_64 arm64"
        if arguments[0] == "/usr/bin/otool":
            # Universal binaries must inspect just the supported slice; an x86 floor is irrelevant.
            assert arguments == ["/usr/bin/otool", "-arch", "arm64", "-l", str(binary)]
            return f"{binary} (architecture arm64):\nLoad command 0\n{metadata}\nLoad command 1\ncmd LC_UUID\n"
        assert arguments == ["/usr/bin/codesign", "--verify", "--deep", "--strict", str(binary)]
        return ""

    monkeypatch.setattr(macos_package, "command", run)
    report = macos_package.verify_signatures(tmp_path)
    assert report["machOCount"] == 1
    assert report["maximumMachOMinimumVersion"] == expected
    assert binary.read_bytes() == original


@pytest.mark.parametrize(
    "metadata",
    [
        "cmd LC_UUID",
        "cmd LC_BUILD_VERSION\n platform MACOS\n sdk 14.0",
        "cmd LC_BUILD_VERSION\n platform MACOS\n minos invalid",
        "cmd LC_BUILD_VERSION\n platform MACOS\n minos 15.0\n sdk 14.0",
        "cmd LC_BUILD_VERSION\n platform MACOS\n minos 14.0.1",
        "cmd LC_BUILD_VERSION\n platform MACOS\n minos 0.0",
        "cmd LC_BUILD_VERSION\n platform IOS\n minos 11.0",
        "cmd LC_BUILD_VERSION\n platform 2\n minos 11.0",
        "cmd LC_BUILD_VERSION\n minos 11.0",
        "cmd LC_BUILD_VERSION\n platform MACOS\n platform 1\n minos 11.0",
        "cmd LC_VERSION_MIN_IPHONEOS\n version 11.0",
        "cmd LC_VERSION_MIN_MACOSX\n version 14.0\n version 13.0",
        "cmd LC_BUILD_VERSION\n platform MACOS\n minos 14.0\nLoad command 1\ncmd LC_VERSION_MIN_MACOSX\n version 13.0",
        "cmd LC_BUILD_VERSION\n platform MACOS\n minos 14.0\nLoad command 1\ncmd LC_BUILD_VERSION\n platform MACOS\n minos 14.0",
    ],
)
def test_invalid_deployment_metadata_rejects_and_names_the_binary(tmp_path, monkeypatch, metadata):
    binary = tmp_path / "nested/library.dylib"
    binary.parent.mkdir()
    binary.write_bytes(b"\xcf\xfa\xed\xfecode")

    def run(arguments):
        if arguments[0] == "/usr/bin/lipo":
            return "arm64"
        if arguments[0] == "/usr/bin/otool":
            return f"Load command 0\n{metadata}\n"
        return ""

    monkeypatch.setattr(macos_package, "command", run)
    with pytest.raises(PublicationError, match=r"nested/library[.]dylib"):
        macos_package.verify_signatures(tmp_path)


def test_signature_report_records_maximum_minimum_across_all_macho_files(tmp_path, monkeypatch):
    for name in ("lower.dylib", "higher.dylib"):
        (tmp_path / name).write_bytes(b"\xcf\xfa\xed\xfecode")

    def run(arguments):
        if arguments[0] == "/usr/bin/lipo":
            return "arm64"
        if arguments[0] == "/usr/bin/otool":
            version = "14.0" if Path(arguments[-1]).name == "higher.dylib" else "11.0"
            return f"Load command 0\ncmd LC_BUILD_VERSION\nplatform MACOS\nminos {version}\nsdk 15.0\n"
        return ""

    monkeypatch.setattr(macos_package, "command", run)
    assert macos_package.verify_signatures(tmp_path)["maximumMachOMinimumVersion"] == "14.0"


@pytest.mark.parametrize("fault", ["first_copy", "second_copy", "report_link", "foreign_archive", "success"])
def test_macos_publication_commits_both_existing_archives_and_report_before_success(tmp_path, monkeypatch, fault):
    """Exercise actual FS publication of saved archives, without native builds/signing."""
    import hashlib
    import json
    import shutil

    source = tmp_path / "already-built"
    source.mkdir()
    archives = []
    for index in range(2):
        content = source / f"payload-{index}.txt"
        content.write_text(f"original archive {index}", encoding="utf-8")
        archive = source / f"candidate-{index}.tar.gz"
        with tarfile.open(archive, "w:gz") as bundle:
            bundle.add(content, arcname=content.name)
        with tarfile.open(archive) as reopened:
            member = reopened.extractfile(content.name)
            assert member is not None and member.read() == content.read_bytes()
        archives.append(archive)
    facts = {
        path.name: {"bytes": path.stat().st_size, "sha256": hashlib.sha256(path.read_bytes()).hexdigest()}
        for path in archives
    }
    originals = {path.name: path.read_bytes() for path in archives}
    output = tmp_path / "candidate"
    report = tmp_path / "receipt.json"
    results = {"schema": "test-native-results", "packages": facts}
    real_copy = shutil.copyfile
    real_link = os.link
    copies = []
    observed_states = []

    def copy_archive(source_path, target_path, **kwargs):
        copies.append(Path(source_path).name)
        if (fault == "first_copy" and len(copies) == 1) or (fault == "second_copy" and len(copies) == 2):
            Path(target_path).write_bytes(Path(source_path).read_bytes()[:11])
            raise OSError("injected partial archive copy")
        return real_copy(source_path, target_path, **kwargs)

    def link_result(source_path, target_path, **kwargs):
        target_path = Path(target_path)
        marker = output / ".docwen-temp-lease.json"
        if target_path in {report, *(output / archive.name for archive in archives)}:
            observed_states.append(json.loads(marker.read_text())["state"])
        if target_path == report and fault == "report_link":
            raise OSError("injected report commit failure")
        if target_path == output / archives[1].name and fault == "foreign_archive":
            target_path.write_bytes(b"foreign archive")
        return real_link(source_path, target_path, **kwargs)

    monkeypatch.setattr(shutil, "copyfile", copy_archive)
    monkeypatch.setattr(os, "link", link_result)
    if fault == "success":
        macos_package._publish_archives(archives, output=output, report=report, results=results)
        assert json.loads(report.read_text()) == results
        assert json.loads((output / ".docwen-temp-lease.json").read_text())["state"] == "completed-success"
        assert not list(tmp_path.glob(".docwen-macos-publish-*"))
        for name, content in originals.items():
            assert (output / name).read_bytes() == content
    else:
        with pytest.raises(OSError):
            macos_package._publish_archives(archives, output=output, report=report, results=results)
        publications = list(tmp_path.glob(".docwen-macos-publish-*"))
        assert len(publications) == 1
        publication = publications[0]
        assert json.loads((publication / ".docwen-temp-lease.json").read_text())["state"] == "retained-failure"
        recovery = json.loads((publication / "publication.json").read_text())
        assert recovery["archives"] == facts
        assert recovery["report"] == {"path": str(report), "content": results}
        assert not report.exists()
        if fault in {"first_copy", "second_copy"}:
            assert not output.exists()
            failed = archives[0 if fault == "first_copy" else 1]
            assert (publication / failed.name).read_bytes() == originals[failed.name][:11]
        else:
            lease = json.loads((output / ".docwen-temp-lease.json").read_text())
            assert lease["state"] == "retained-failure"
            assert lease["artifacts"] == facts
            assert lease["publicationRoot"] == str(publication)
            assert (output / archives[0].name).read_bytes() == originals[archives[0].name]
            if fault == "foreign_archive":
                assert (output / archives[1].name).read_bytes() == b"foreign archive"
            else:
                assert (output / archives[1].name).read_bytes() == originals[archives[1].name]
    assert all(state == "active" for state in observed_states)
    assert {path.name: path.read_bytes() for path in archives} == originals
    assert len(copies) == (1 if fault == "first_copy" else 2)


@pytest.mark.parametrize("existing", ["output", "report"])
def test_macos_publication_refuses_unknown_preexisting_outputs(tmp_path, existing):
    output = tmp_path / "candidate"
    report = tmp_path / "receipt.json"
    if existing == "output":
        output.mkdir()
        foreign = output / "private"
    else:
        foreign = report
    foreign.write_bytes(b"original owner")
    with pytest.raises(PublicationError, match="must be new"):
        macos_package._publish_archives([], output=output, report=report, results={})
    assert foreign.read_bytes() == b"original owner"
    assert not list(tmp_path.glob(".docwen-macos-publish-*"))


@pytest.mark.parametrize("partial_marker", [False, True])
def test_macos_candidate_marker_initialization_preserves_reserved_identity(tmp_path, monkeypatch, partial_marker):
    import json

    output = tmp_path / "candidate"
    real_record = macos_package.atomic_record

    def record(path, content):
        if path == output / ".docwen-temp-lease.json":
            if partial_marker:
                path.write_bytes(content[:19])
            raise OSError("injected candidate marker initialization")
        return real_record(path, content)

    monkeypatch.setattr(macos_package, "atomic_record", record)
    with pytest.raises(OSError, match="candidate marker initialization"):
        macos_package._publish_archives([], output=output, report=tmp_path / "report.json", results={})
    [publication] = list(tmp_path.glob(".docwen-macos-publish-*"))
    assert json.loads((publication / ".docwen-temp-lease.json").read_text())["state"] == "retained-failure"
    reservation = json.loads((publication / "candidate-identity.json").read_text())
    assert reservation["root"] == str(output)
    assert reservation["owner"] == "docwen.release.macos"
    if partial_marker:
        assert reservation["rootIdentity"] == {"device": output.stat().st_dev, "inode": output.stat().st_ino}
        assert len((output / ".docwen-temp-lease.json").read_bytes()) == 19
    else:
        assert not output.exists()
    assert not (tmp_path / "report.json").exists()


@pytest.mark.parametrize("recovery_failure", [False, True])
def test_macos_terminal_record_failure_keeps_completed_publication_fact(tmp_path, monkeypatch, recovery_failure):
    import json

    from tools import run_lease

    source = payload(tmp_path)
    archives = [tmp_path / "cli.tar.gz", tmp_path / "gui.tar.gz"]
    for archive in archives:
        archive_payload(source, archive, top="App", epoch=1)
        with tarfile.open(archive) as reopened:
            member = reopened.extractfile("App/文档.txt")
            assert member is not None and member.read() == b"Unicode content"
    original = {archive.name: archive.read_bytes() for archive in archives}
    output = tmp_path / "candidate"
    report = tmp_path / "receipt.json"
    real_dump = run_lease.json.dump
    real_fsync = run_lease.os.fsync
    real_open = Path.open
    failing_fd = None

    def dump(value, stream, *args, **kwargs):
        nonlocal failing_fd
        result = real_dump(value, stream, *args, **kwargs)
        if value.get("root") == str(output) and value.get("state") == "completed-success":
            failing_fd = stream.fileno()
        return result

    def fsync(fd):
        nonlocal failing_fd
        if fd == failing_fd:
            failing_fd = None
            raise OSError("injected candidate terminal fsync failure")
        return real_fsync(fd)

    def open_path(path, *args, **kwargs):
        if recovery_failure and path == output.with_name("candidate.recovery.json"):
            raise OSError("injected external recovery failure")
        return real_open(path, *args, **kwargs)

    monkeypatch.setattr(run_lease.json, "dump", dump)
    monkeypatch.setattr(run_lease.os, "fsync", fsync)
    monkeypatch.setattr(Path, "open", open_path)
    with pytest.raises(OSError, match="terminal fsync") as caught:
        macos_package._publish_archives(archives, output=output, report=report, results={"complete": True})
    assert json.loads(report.read_text()) == {"complete": True}
    assert {name: (output / name).read_bytes() for name in original} == original
    # The primary failure is evidence failure after publication, not a failed
    # archive build. The failed .next retains the true completed outcome.
    pending = json.loads((output / ".docwen-temp-lease.next").read_text())
    assert pending["state"] == pending["outcome"] == "completed-success"
    assert json.loads((output / ".docwen-temp-lease.json").read_text())["state"] == "active"
    if recovery_failure:
        assert any("external recovery failure" in note for note in caught.value.__notes__)
    else:
        recovery = json.loads(output.with_name("candidate.recovery.json").read_text())
        assert recovery["phase"] == "terminal-recording"
        assert recovery["knownOutcome"] == recovery["lease"]["outcome"] == "completed-success"
        assert recovery["executionError"] is None
