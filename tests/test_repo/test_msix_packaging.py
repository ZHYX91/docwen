"""Microsoft Store MSIX packaging contract tests."""

from __future__ import annotations

import json
import os
import struct
import zipfile
from pathlib import Path
from xml.etree import ElementTree

import pytest
from PIL import Image
from scripts.release import build_msix

pytestmark = pytest.mark.contract

_REPO_ROOT = Path(__file__).resolve().parents[2]
_CONFIG_PATH = _REPO_ROOT / "release" / "windows-store-msix.v1.json"
_FOUNDATION = "http://schemas.microsoft.com/appx/manifest/foundation/windows10"
_UAP5 = "http://schemas.microsoft.com/appx/manifest/uap/windows10/5"
_RESCAP = "http://schemas.microsoft.com/appx/manifest/foundation/windows10/restrictedcapabilities"


def _fake_payload(path: Path) -> Path:
    path.mkdir()
    (path / "DocWen.exe").write_bytes(b"not-a-real-pe")
    (path / "DocWenCLI.exe").write_bytes(b"not-a-real-pe")
    assets = path / "assets"
    assets.mkdir()
    Image.new("RGBA", (256, 256), (38, 111, 227, 255)).save(assets / "icon.png")
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
    namespaces = {"f": _FOUNDATION, "uap5": _UAP5, "rescap": _RESCAP}

    identity = manifest.find("f:Identity", namespaces)
    assert identity is not None
    assert identity.attrib == {
        "Name": "ZHYX.DocWen",
        "Publisher": "CN=9E46E7F1-F057-4B88-BF71-7C9CB77AF9C6",
        "Version": "1.0.1.0",
        "ProcessorArchitecture": "x64",
    }
    target = manifest.find("f:Dependencies/f:TargetDeviceFamily", namespaces)
    assert target is not None and target.attrib["Name"] == "Windows.Desktop"
    capability = manifest.find("f:Capabilities/rescap:Capability", namespaces)
    assert capability is not None and capability.attrib["Name"] == "runFullTrust"
    alias = manifest.find(".//uap5:ExecutionAlias", namespaces)
    assert alias is not None and alias.attrib["Alias"] == "docwen.exe"


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
    for name, expected_size in build_msix._ASSET_SIZES.items():  # pyright: ignore[reportPrivateUsage]
        with Image.open(staging / "assets" / "msix" / name) as image:
            assert image.size == expected_size


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


@pytest.mark.windows_only
def test_makeappx_accepts_generated_layout_when_windows_sdk_is_available(tmp_path: Path) -> None:
    if os.name != "nt":
        pytest.skip("Windows-only MakeAppx contract")
    try:
        makeappx = build_msix.find_makeappx()
    except build_msix.MsixBuildError:
        pytest.skip("Windows SDK MakeAppx is unavailable")

    config = build_msix.read_config(_CONFIG_PATH)
    output = tmp_path / str(config["assetName"])
    metadata = build_msix.build_msix(
        payload_root=_fake_payload(tmp_path / "payload"),
        output=output,
        work_root=tmp_path / "work",
        config=config,
        makeappx=makeappx,
    )

    assert metadata["assetName"] == "DocWen-windows-x64.msix"
    assert len(str(metadata["sha256"])) == 64
    assert len(str(metadata["contentSha256"])) == 64
    assert int(metadata["entryCount"]) > 5
    with zipfile.ZipFile(output) as package:
        names = set(package.namelist())
        assert "AppxManifest.xml" in names
        assert "assets/msix/StoreLogo.png" in names
        assert "DocWen.exe" in names
        assert "DocWenCLI.exe" in names


def test_checked_in_store_config_is_stable_json() -> None:
    config = json.loads(_CONFIG_PATH.read_text(encoding="utf-8"))
    assert config["storeId"] == "9NR2211SJH97"
    assert config["sourceVersion"] == "0.9.0"


def test_package_content_identity_ignores_zip_timestamps(tmp_path: Path) -> None:
    first = tmp_path / "first.msix"
    second = tmp_path / "second.msix"
    for path, timestamp in ((first, (2025, 1, 2, 3, 4, 6)), (second, (2026, 2, 3, 4, 5, 8))):
        with zipfile.ZipFile(path, "w") as package:
            entry = zipfile.ZipInfo("payload.txt", date_time=timestamp)
            package.writestr(entry, b"same payload")

    assert first.read_bytes() != second.read_bytes()
    assert build_msix.package_content_identity(first) == build_msix.package_content_identity(second)
