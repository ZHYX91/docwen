"""Microsoft Store MSIX packaging contract tests."""

from __future__ import annotations

import json
import os
import stat
import struct
import subprocess
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree

import pytest
from PIL import Image
from scripts.release import build_msix
from scripts.release.msix_candidate import extract_candidate
from scripts.release.publication_contract import PublicationError, file_identity

pytestmark = pytest.mark.contract

_REPO_ROOT = Path(__file__).resolve().parents[2]
_CONFIG_PATH = _REPO_ROOT / "release" / "windows-store-msix.v1.json"
_FOUNDATION = "http://schemas.microsoft.com/appx/manifest/foundation/windows10"
_UAP = "http://schemas.microsoft.com/appx/manifest/uap/windows10"
_UAP5 = "http://schemas.microsoft.com/appx/manifest/uap/windows10/5"
_RESCAP = "http://schemas.microsoft.com/appx/manifest/foundation/windows10/restrictedcapabilities"


def _fake_payload(path: Path) -> Path:
    path.mkdir()
    (path / "DocWen.exe").write_bytes(b"not-a-real-pe")
    (path / "DocWenCLI.exe").write_bytes(b"not-a-real-pe")
    assets = path / "assets"
    assets.mkdir()
    Image.new("RGBA", (256, 256), (38, 111, 227, 255)).save(assets / "icon.png")
    (assets / "icon.svg").write_text(
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">'
        '<rect width="100" height="100" fill="#E5484D" /></svg>',
        encoding="utf-8",
    )
    docx_templates = path / "_internal" / "docx" / "templates"
    expanded_template = docx_templates / "default-docx-template"
    expanded_template.mkdir(parents=True)
    (expanded_template / "[Content_Types].xml").write_text("reserved OPC part", encoding="utf-8")
    (docx_templates / "default.docx").write_bytes(b"runtime-template")
    return path


def _fake_pe_with_certificate_directory(
    path: Path,
    *,
    certificate_offset: int,
    certificate_size: int,
    file_size: int = 1024,
) -> Path:
    image = bytearray(file_size)
    image[:2] = b"MZ"
    pe_offset = 0x80
    struct.pack_into("<I", image, 0x3C, pe_offset)
    image[pe_offset : pe_offset + 4] = b"PE\0\0"
    optional_size = 240
    struct.pack_into("<H", image, pe_offset + 4 + 16, optional_size)
    optional_offset = pe_offset + 24
    struct.pack_into("<H", image, optional_offset, 0x20B)
    struct.pack_into("<I", image, optional_offset + 108, 16)
    struct.pack_into(
        "<II",
        image,
        optional_offset + 112 + (4 * 8),
        certificate_offset,
        certificate_size,
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(image)
    return path


def test_store_config_and_manifest_bind_partner_center_identity() -> None:
    config = build_msix.read_config(_CONFIG_PATH)
    manifest = ElementTree.fromstring(build_msix.render_manifest(config))
    namespaces = {"f": _FOUNDATION, "uap": _UAP, "uap5": _UAP5, "rescap": _RESCAP}

    identity = manifest.find("f:Identity", namespaces)
    assert identity is not None
    assert identity.attrib == {
        "Name": "ZHYX.DocWen",
        "Publisher": "CN=9E46E7F1-F057-4B88-BF71-7C9CB77AF9C6",
        "Version": "1.0.7.0",
        "ProcessorArchitecture": "x64",
    }
    target = manifest.find("f:Dependencies/f:TargetDeviceFamily", namespaces)
    assert target is not None and target.attrib["Name"] == "Windows.Desktop"
    capability = manifest.find("f:Capabilities/rescap:Capability", namespaces)
    assert capability is not None and capability.attrib["Name"] == "runFullTrust"
    alias = manifest.find(".//uap5:ExecutionAlias", namespaces)
    assert alias is not None and alias.attrib["Alias"] == "docwen.exe"
    default_tile = manifest.find(".//f:Application/uap:VisualElements/uap:DefaultTile", namespaces)
    assert default_tile is not None and default_tile.attrib["ShortName"] == "DocWen"
    show_name = default_tile.find("uap:ShowNameOnTiles/uap:ShowOn", namespaces)
    assert show_name is not None and show_name.attrib["Tile"] == "square150x150Logo"


@pytest.mark.parametrize("version", ["0.9.0.0", "1.0.0.1", "1.0.0", "65536.0.0.0"])
def test_store_package_version_rejects_values_partner_center_will_reject(version: str) -> None:
    with pytest.raises(build_msix.MsixBuildError):
        build_msix.validate_package_version(version)


def test_prepare_layout_keeps_payload_assets_and_generates_required_logos(tmp_path: Path) -> None:
    config = build_msix.read_config(_CONFIG_PATH)
    payload = _fake_payload(tmp_path / "payload")
    staging = tmp_path / "staging"

    build_msix.prepare_layout(payload, staging, config)

    assert (staging / "DocWen.exe").read_bytes() == b"not-a-real-pe"
    assert (staging / "AppxManifest.xml").is_file()
    assert not (staging / "_internal" / "docx" / "templates" / "default-docx-template").exists()
    assert (staging / "_internal" / "docx" / "templates" / "default.docx").is_file()
    for name, (expected_size, fill_ratio) in build_msix._ASSET_SPECS.items():  # pyright: ignore[reportPrivateUsage]
        with Image.open(staging / "assets" / "msix" / name) as image:
            assert image.size == expected_size
            alpha_bounds = image.getchannel("A").getbbox()
            assert alpha_bounds is not None
            assert alpha_bounds[2] - alpha_bounds[0] == round(expected_size[0] * fill_ratio)
            assert alpha_bounds[3] - alpha_bounds[1] == round(expected_size[1] * fill_ratio)
            assert image.getpixel((image.width // 2, image.height // 2))[:3] == (229, 72, 77)


def test_store_png_halfway_alpha_rounding_preserves_colour_and_transparency() -> None:
    # The green channel is exactly 110.5 after unpremultiplication; CI hosts
    # previously disagreed between 110 and 111 for these real logo pixels.
    pixels = bytes((33, 78, 166, 180, 0, 0, 0, 0, 47, 111, 235, 255))
    from scripts.icon_pixels import unpremultiply_rgba

    assert unpremultiply_rgba(pixels) == bytes((47, 111, 235, 180, 0, 0, 0, 0, 47, 111, 235, 255))


def test_store_icons_are_identical_across_optional_qt_cpu_features(tmp_path: Path) -> None:
    if os.name != "nt":
        pytest.skip("Windows Store rendering CPU regression")
    command = (
        "from pathlib import Path; import sys; "
        "from scripts.release.build_msix import _write_assets; "
        "_write_assets(Path(sys.argv[1]))"
    )
    roots = [tmp_path / "native-cpu", tmp_path / "baseline-cpu"]
    for root, features in zip(roots, ("", "avx2 avx512f"), strict=True):
        env = os.environ.copy()
        env["QT_NO_CPU_FEATURE"] = features
        subprocess.run([sys.executable, "-c", command, str(root)], cwd=_REPO_ROOT, env=env, check=True)
    for name in build_msix._ASSET_SPECS:  # pyright: ignore[reportPrivateUsage]
        assert (roots[0] / "assets/msix" / name).read_bytes() == (roots[1] / "assets/msix" / name).read_bytes(), name


def test_prepare_layout_clears_certificate_pointer_when_signature_blob_was_stripped(tmp_path: Path) -> None:
    config = build_msix.read_config(_CONFIG_PATH)
    payload = _fake_payload(tmp_path / "payload")
    stripped = _fake_pe_with_certificate_directory(
        payload / "_internal" / "tcl86t.dll",
        certificate_offset=1024,
        certificate_size=7408,
    )
    staging = tmp_path / "staging"

    sanitized = build_msix.prepare_layout(payload, staging, config)

    assert sanitized == ("_internal/tcl86t.dll",)
    staged = (staging / stripped.relative_to(payload)).read_bytes()
    security_entry_offset = 0x80 + 24 + 112 + (4 * 8)
    assert struct.unpack_from("<II", staged, security_entry_offset) == (0, 0)


def test_prepare_layout_rejects_partially_truncated_certificate_table(tmp_path: Path) -> None:
    config = build_msix.read_config(_CONFIG_PATH)
    payload = _fake_payload(tmp_path / "payload")
    _fake_pe_with_certificate_directory(
        payload / "_internal" / "broken.dll",
        certificate_offset=1016,
        certificate_size=16,
    )

    with pytest.raises(
        build_msix.MsixBuildError,
        match=r"msix_pe_certificate_table_corrupt:_internal/broken\.dll",
    ):
        build_msix.prepare_layout(payload, tmp_path / "staging", config)


def _portable_candidate(tmp_path: Path, payload: Path) -> tuple[Path, Path]:
    archive = tmp_path / "DocWen-windows-x64.zip"
    with zipfile.ZipFile(archive, "w") as package:
        for path in sorted(payload.rglob("*")):
            if path.is_file():
                package.write(path, path.relative_to(payload).as_posix())
    return archive, _inspection_receipt(tmp_path, archive)


def _inspection_receipt(tmp_path: Path, archive: Path) -> Path:
    receipt = tmp_path / "inspection.json"
    receipt.write_text(
        json.dumps(
            {
                "stage": "candidate-verified",
                "provenance": "verified",
                "repository": "ZHYX91/docwen",
                "version": "0.13.1",
                "sourceCommit": "a" * 40,
                "manifestSha256": "b" * 64,
                "artifactId": 20,
                "artifactDigest": "sha256:" + "c" * 64,
                "origin": {"sourceRef": "refs/heads/release/candidate", "runId": 10, "runAttempt": 1},
                "assets": {"DocWen-windows-x64.zip": file_identity(archive)},
            }
        ),
        encoding="utf-8",
    )
    return receipt


@pytest.mark.windows_only
@pytest.mark.parametrize("from_archive", [False, True])
def test_makeappx_accepts_generated_layout_when_windows_sdk_is_available(tmp_path: Path, from_archive: bool) -> None:
    if os.name != "nt":
        pytest.skip("Windows-only MakeAppx contract")
    try:
        makeappx = build_msix.find_makeappx()
    except build_msix.MsixBuildError:
        pytest.skip("Windows SDK MakeAppx is unavailable")

    config = build_msix.read_config(_CONFIG_PATH)
    output = tmp_path / str(config["assetName"])
    payload = _fake_payload(tmp_path / "payload")
    archive, receipt = _portable_candidate(tmp_path, payload)
    metadata = build_msix.build_msix(
        payload_root=None if from_archive else payload,
        portable_zip=archive if from_archive else None,
        candidate_receipt=receipt if from_archive else None,
        output=output,
        work_root=tmp_path / "work",
        config=config,
        makeappx=makeappx,
    )

    assert metadata["assetName"] == "DocWen-windows-x64.msix"
    assert len(str(metadata["sha256"])) == 64
    assert len(str(metadata["contentSha256"])) == 64
    assert int(metadata["entryCount"]) > 5
    assert not (tmp_path / "work").exists()
    if from_archive:
        assert metadata["sourceCandidate"]["portableArchive"] == file_identity(archive)
        assert metadata["sourceCandidate"]["artifactId"] == 20
    else:
        assert "sourceCandidate" not in metadata
    with zipfile.ZipFile(output) as package:
        names = set(package.namelist())
        assert "AppxManifest.xml" in names
        assert "assets/msix/StoreLogo.png" in names
        assert "DocWen.exe" in names
        assert "DocWenCLI.exe" in names


@pytest.mark.parametrize(
    "name,mode",
    [
        ("../escaped.txt", stat.S_IFREG),
        ("/absolute.txt", stat.S_IFREG),
        ("folder\\escaped.txt", stat.S_IFREG),
        ("C:escaped.txt", stat.S_IFREG),
        ("folder./file.txt", stat.S_IFREG),
        ("NUL.txt", stat.S_IFREG),
        ("DOCWEN.exe", stat.S_IFREG),
        ("linked.txt", stat.S_IFLNK),
    ],
)
def test_portable_candidate_rejects_unsafe_entries_before_extraction(tmp_path: Path, name: str, mode: int) -> None:
    archive = tmp_path / "candidate.zip"
    with zipfile.ZipFile(archive, "w") as package:
        for filename, entry_mode in [("DocWen.exe", stat.S_IFREG), ("DocWenCLI.exe", stat.S_IFREG), (name, mode)]:
            entry = zipfile.ZipInfo(filename)
            entry.filename = filename  # Preserve raw separators instead of ZipInfo's Windows normalization.
            entry.external_attr = (entry_mode | 0o644) << 16
            package.writestr(entry, b"test")
    receipt = _inspection_receipt(tmp_path, archive)
    output = tmp_path / "extracted"
    with pytest.raises(PublicationError):
        extract_candidate(archive, receipt, output, version="0.13.1")
    assert not output.exists()
    assert not (tmp_path / "escaped.txt").exists()


@pytest.mark.parametrize("failure", ["changed-archive", "uninspected", "wrong-version"])
def test_portable_candidate_requires_inspected_bytes_and_version(tmp_path: Path, failure: str) -> None:
    archive, receipt = _portable_candidate(tmp_path, _fake_payload(tmp_path / "payload"))
    version = "0.13.1"
    if failure == "changed-archive":
        with archive.open("ab") as stream:
            stream.write(b"changed")
    elif failure == "uninspected":
        record = json.loads(receipt.read_text(encoding="utf-8"))
        record["stage"] = "prepared"
        receipt.write_text(json.dumps(record), encoding="utf-8")
    else:
        version = "99.0.0"
    output = tmp_path / "extracted"
    with pytest.raises(PublicationError):
        extract_candidate(archive, receipt, output, version=version)
    assert not output.exists()


def test_msix_never_replaces_an_existing_work_directory(tmp_path: Path) -> None:
    work = tmp_path / "work"
    work.mkdir()
    original = work / "user-data.txt"
    original.write_bytes(b"keep me")
    with pytest.raises(build_msix.MsixBuildError, match="msix_work_root_already_exists"):
        build_msix.build_msix(
            payload_root=_fake_payload(tmp_path / "payload"),
            output=tmp_path / "output.msix",
            work_root=work,
            config=build_msix.read_config(_CONFIG_PATH),
            makeappx=tmp_path / "unused.exe",
        )
    assert original.read_bytes() == b"keep me"
    assert set(work.iterdir()) == {original}


def test_msix_build_failure_retains_an_owned_run_and_does_not_overwrite_output(tmp_path: Path) -> None:
    work = tmp_path / "work"
    output = tmp_path / "existing.msix"
    output.write_bytes(b"previous package")
    with pytest.raises(build_msix.MsixBuildError, match="msix_output_must_be_new_outside_work"):
        build_msix.build_msix(
            payload_root=_fake_payload(tmp_path / "payload"),
            output=output,
            work_root=work,
            config=build_msix.read_config(_CONFIG_PATH),
            makeappx=tmp_path / "unused.exe",
        )
    assert output.read_bytes() == b"previous package"
    lease = json.loads((work / ".docwen-temp-lease.json").read_text(encoding="utf-8"))
    assert lease["owner"] == "docwen.release.msix"
    assert lease["state"] == "retained-failure"


def test_checked_in_store_config_is_stable_json() -> None:
    config = json.loads(_CONFIG_PATH.read_text(encoding="utf-8"))
    assert config["storeId"] == "9NR2211SJH97"
    assert config["sourceVersion"] == "0.13.1"


def test_package_content_identity_ignores_zip_timestamps(tmp_path: Path) -> None:
    first = tmp_path / "first.msix"
    second = tmp_path / "second.msix"
    for path, timestamp in ((first, (2025, 1, 2, 3, 4, 6)), (second, (2026, 2, 3, 4, 5, 8))):
        with zipfile.ZipFile(path, "w") as package:
            entry = zipfile.ZipInfo("payload.txt", date_time=timestamp)
            package.writestr(entry, b"same payload")

    assert first.read_bytes() != second.read_bytes()
    assert build_msix.package_content_identity(first) == build_msix.package_content_identity(second)
