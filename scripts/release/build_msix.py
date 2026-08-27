#!/usr/bin/env python3
"""Build the unsigned Microsoft Store MSIX from a verified Windows payload.

The Store signs an accepted package. This script deliberately does not create
or manage a local signing identity; it binds the package to the exact identity
reserved in Partner Center and emits deterministic submission metadata.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import struct
import subprocess
import sys
import zipfile
from collections.abc import Mapping, Sequence
from pathlib import Path, PurePosixPath
from typing import Any, cast
from xml.sax.saxutils import escape, quoteattr

from PIL import Image

_REPO_ROOT = Path(__file__).resolve().parents[2]
_SEMVER = re.compile(r"^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$")
_FOUR_PART_VERSION = re.compile(r"^\d+\.\d+\.\d+\.\d+$")
_ASSET_SIZES = {
    "StoreLogo.png": (50, 50),
    "Square44x44Logo.png": (44, 44),
    "Square150x150Logo.png": (150, 150),
}


class MsixBuildError(RuntimeError):
    """A stable, fail-closed MSIX build error."""


def _require(condition: bool, message: str) -> None:
    if not condition:
        raise MsixBuildError(message)


def _as_object(value: object, label: str) -> dict[str, Any]:
    _require(isinstance(value, dict), f"msix_config_{label}_invalid")
    return cast(dict[str, Any], value)


def _as_text(value: object, label: str) -> str:
    _require(isinstance(value, str) and bool(value.strip()), f"msix_config_{label}_invalid")
    return cast(str, value)


def validate_package_version(value: str) -> tuple[int, int, int, int]:
    """Validate the Store's four-part package version contract."""

    _require(_FOUR_PART_VERSION.fullmatch(value) is not None, "msix_package_version_format")
    parts = tuple(int(part) for part in value.split("."))
    _require(len(parts) == 4, "msix_package_version_format")
    _require(all(0 <= part <= 65535 for part in parts), "msix_package_version_range")
    _require(parts[0] != 0, "msix_package_version_major_zero")
    _require(parts[3] == 0, "msix_package_version_revision_nonzero")
    return parts  # type: ignore[return-value]


def read_config(path: Path) -> dict[str, Any]:
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise MsixBuildError("msix_config_unreadable") from exc

    config = _as_object(raw, "root")
    expected_keys = {
        "schemaVersion",
        "storeId",
        "sourceVersion",
        "packageVersion",
        "identity",
        "application",
        "target",
        "languages",
        "excludedPayloadPaths",
        "assetName",
        "reproducibilityEpoch",
    }
    _require(set(config) == expected_keys, "msix_config_keys")
    _require(config["schemaVersion"] == 1, "msix_config_schema_version")
    _require(_as_text(config["storeId"], "store_id") == "9NR2211SJH97", "msix_config_store_id")

    source_version = _as_text(config["sourceVersion"], "source_version")
    _require(_SEMVER.fullmatch(source_version) is not None, "msix_config_source_version")
    validate_package_version(_as_text(config["packageVersion"], "package_version"))

    identity = _as_object(config["identity"], "identity")
    _require(set(identity) == {"name", "publisher", "publisherDisplayName"}, "msix_config_identity_keys")
    _require(_as_text(identity["name"], "identity_name") == "ZHYX.DocWen", "msix_config_identity_name")
    _require(
        _as_text(identity["publisher"], "identity_publisher") == "CN=9E46E7F1-F057-4B88-BF71-7C9CB77AF9C6",
        "msix_config_identity_publisher",
    )
    _as_text(identity["publisherDisplayName"], "publisher_display_name")

    application = _as_object(config["application"], "application")
    _require(
        set(application) == {"id", "displayName", "description", "executable", "cliExecutable", "cliAlias"},
        "msix_config_application_keys",
    )
    for key in application:
        _as_text(application[key], f"application_{key}")
    _require(
        str(application["executable"]).casefold() != str(application["cliExecutable"]).casefold(), "msix_exes_distinct"
    )
    _require(str(application["cliAlias"]).casefold().endswith(".exe"), "msix_cli_alias_extension")

    target = _as_object(config["target"], "target")
    _require(set(target) == {"architecture", "minVersion", "maxVersionTested"}, "msix_config_target_keys")
    _require(target["architecture"] == "x64", "msix_config_architecture")
    _as_text(target["minVersion"], "min_version")
    _as_text(target["maxVersionTested"], "max_version_tested")

    languages = config["languages"]
    _require(isinstance(languages, list) and languages == ["zh-CN", "en-US"], "msix_config_languages")
    excluded = config["excludedPayloadPaths"]
    _require(
        excluded == ["_internal/docx/templates/default-docx-template"],
        "msix_config_excluded_payload_paths",
    )
    _require(_as_text(config["assetName"], "asset_name").endswith(".msix"), "msix_config_asset_name")
    epoch = config["reproducibilityEpoch"]
    _require(isinstance(epoch, int) and epoch > 0, "msix_config_reproducibility_epoch")
    return config


def render_manifest(config: Mapping[str, Any]) -> str:
    identity = config["identity"]
    application = config["application"]
    target = config["target"]
    resources = "\n".join(f"    <Resource Language={quoteattr(str(language))} />" for language in config["languages"])
    values = {
        "identity_name": quoteattr(str(identity["name"])),
        "publisher": quoteattr(str(identity["publisher"])),
        "package_version": quoteattr(str(config["packageVersion"])),
        "architecture": quoteattr(str(target["architecture"])),
        "display_name": escape(str(application["displayName"])),
        "publisher_display_name": escape(str(identity["publisherDisplayName"])),
        "min_version": quoteattr(str(target["minVersion"])),
        "max_version": quoteattr(str(target["maxVersionTested"])),
        "application_id": quoteattr(str(application["id"])),
        "executable": quoteattr(str(application["executable"])),
        "description": quoteattr(str(application["description"])),
        "cli_executable": quoteattr(str(application["cliExecutable"])),
        "cli_alias": quoteattr(str(application["cliAlias"])),
    }
    return f"""<?xml version="1.0" encoding="utf-8"?>
<Package
  xmlns="http://schemas.microsoft.com/appx/manifest/foundation/windows10"
  xmlns:uap="http://schemas.microsoft.com/appx/manifest/uap/windows10"
  xmlns:uap5="http://schemas.microsoft.com/appx/manifest/uap/windows10/5"
  xmlns:desktop4="http://schemas.microsoft.com/appx/manifest/desktop/windows10/4"
  xmlns:rescap="http://schemas.microsoft.com/appx/manifest/foundation/windows10/restrictedcapabilities"
  IgnorableNamespaces="uap uap5 desktop4 rescap">
  <Identity Name={values["identity_name"]} Publisher={values["publisher"]}
    Version={values["package_version"]} ProcessorArchitecture={values["architecture"]} />
  <Properties>
    <DisplayName>{values["display_name"]}</DisplayName>
    <PublisherDisplayName>{values["publisher_display_name"]}</PublisherDisplayName>
    <Logo>assets\\msix\\StoreLogo.png</Logo>
  </Properties>
  <Resources>
{resources}
  </Resources>
  <Dependencies>
    <TargetDeviceFamily Name="Windows.Desktop" MinVersion={values["min_version"]}
      MaxVersionTested={values["max_version"]} />
  </Dependencies>
  <Applications>
    <Application Id={values["application_id"]} Executable={values["executable"]}
      EntryPoint="Windows.FullTrustApplication" desktop4:SupportsMultipleInstances="true">
      <uap:VisualElements DisplayName={quoteattr(str(application["displayName"]))}
        Description={values["description"]} BackgroundColor="transparent"
        Square150x150Logo="assets\\msix\\Square150x150Logo.png"
        Square44x44Logo="assets\\msix\\Square44x44Logo.png" />
      <Extensions>
        <uap5:Extension Category="windows.appExecutionAlias" Executable={values["cli_executable"]}
          EntryPoint="Windows.FullTrustApplication">
          <uap5:AppExecutionAlias desktop4:Subsystem="console">
            <uap5:ExecutionAlias Alias={values["cli_alias"]} />
          </uap5:AppExecutionAlias>
        </uap5:Extension>
      </Extensions>
    </Application>
  </Applications>
  <Capabilities>
    <rescap:Capability Name="runFullTrust" />
  </Capabilities>
</Package>
"""


def _copy_payload(payload_root: Path, staging_root: Path, config: Mapping[str, Any]) -> None:
    _require(payload_root.is_dir(), "msix_payload_root_missing")
    _require(not (payload_root / "AppxManifest.xml").exists(), "msix_payload_manifest_preexisting")
    application = config["application"]
    for executable_key in ("executable", "cliExecutable"):
        _require((payload_root / application[executable_key]).is_file(), f"msix_payload_{executable_key}_missing")

    for source in payload_root.rglob("*"):
        _require(not source.is_symlink(), "msix_payload_links_forbidden")
    shutil.copytree(payload_root, staging_root)

    for relative_text in config["excludedPayloadPaths"]:
        relative = PurePosixPath(relative_text)
        _require(not relative.is_absolute() and ".." not in relative.parts, "msix_excluded_payload_path_invalid")
        excluded_path = staging_root.joinpath(*relative.parts)
        _require(excluded_path.is_dir() and not excluded_path.is_symlink(), "msix_excluded_payload_path_missing")
        shutil.rmtree(excluded_path)
    _require(
        (staging_root / "_internal" / "docx" / "templates" / "default.docx").is_file(),
        "msix_default_docx_missing",
    )


def _write_assets(staging_root: Path) -> None:
    candidates = (staging_root / "assets" / "icon.png", _REPO_ROOT / "assets" / "icon.png")
    icon_path = next((path for path in candidates if path.is_file()), None)
    _require(icon_path is not None, "msix_icon_source_missing")
    icon_path = cast(Path, icon_path)
    output_root = staging_root / "assets" / "msix"
    output_root.mkdir(parents=True, exist_ok=True)

    with Image.open(icon_path) as source:
        icon = source.convert("RGBA")
        for name, size in _ASSET_SIZES.items():
            resized = icon.resize(size, Image.Resampling.LANCZOS)
            resized.save(output_root / name, format="PNG", optimize=False, compress_level=9)


def _sanitize_stripped_pe_certificates(staging_root: Path) -> tuple[str, ...]:
    """Clear certificate table pointers whose signature bytes were stripped.

    Some standalone CPython distributions retain a PE security-directory entry
    after removing the certificate blob from Tcl/Tk DLLs. Windows can execute
    those DLLs, but SignTool refuses to sign any MSIX containing them with
    ``0x800700C1``. Clearing an entry is safe only when it starts exactly at EOF;
    any other malformed certificate table fails closed.
    """

    sanitized: list[str] = []
    for path in sorted(candidate for candidate in staging_root.rglob("*") if candidate.is_file()):
        file_size = path.stat().st_size
        if file_size < 64:
            continue
        with path.open("rb") as stream:
            if stream.read(2) != b"MZ":
                continue
            stream.seek(0x3C)
            pe_offset_bytes = stream.read(4)
            _require(len(pe_offset_bytes) == 4, "msix_pe_header_truncated")
            pe_offset = struct.unpack("<I", pe_offset_bytes)[0]
            _require(pe_offset + 24 <= file_size, "msix_pe_header_invalid")
            stream.seek(pe_offset)
            _require(stream.read(4) == b"PE\0\0", "msix_pe_signature_invalid")
            coff_header = stream.read(20)
            _require(len(coff_header) == 20, "msix_pe_header_truncated")
            optional_size = struct.unpack_from("<H", coff_header, 16)[0]
            optional_offset = pe_offset + 24
            _require(optional_offset + optional_size <= file_size, "msix_pe_optional_header_invalid")
            stream.seek(optional_offset)
            magic_bytes = stream.read(2)
            _require(len(magic_bytes) == 2, "msix_pe_optional_header_truncated")
            magic = struct.unpack("<H", magic_bytes)[0]
            if magic == 0x10B:
                data_directory_offset = 96
                directory_count_offset = 92
            elif magic == 0x20B:
                data_directory_offset = 112
                directory_count_offset = 108
            else:
                raise MsixBuildError("msix_pe_optional_header_magic")
            _require(
                optional_size >= directory_count_offset + 4,
                "msix_pe_optional_header_directories_missing",
            )
            stream.seek(optional_offset + directory_count_offset)
            directory_count = struct.unpack("<I", stream.read(4))[0]
            if directory_count <= 4:
                continue
            security_entry_offset = optional_offset + data_directory_offset + (4 * 8)
            _require(
                security_entry_offset + 8 <= optional_offset + optional_size,
                "msix_pe_security_directory_missing",
            )
            stream.seek(security_entry_offset)
            certificate_offset, certificate_size = struct.unpack("<II", stream.read(8))

        if certificate_offset == 0 and certificate_size == 0:
            continue
        relative = path.relative_to(staging_root).as_posix()
        if certificate_offset == file_size and certificate_size > 0:
            with path.open("r+b") as stream:
                stream.seek(security_entry_offset)
                stream.write(b"\0" * 8)
            sanitized.append(relative)
            continue
        _require(
            certificate_offset > 0 and certificate_size >= 8 and certificate_offset + certificate_size <= file_size,
            f"msix_pe_certificate_table_corrupt:{relative}",
        )
    return tuple(sanitized)


def prepare_layout(payload_root: Path, staging_root: Path, config: Mapping[str, Any]) -> tuple[str, ...]:
    _copy_payload(payload_root, staging_root, config)
    sanitized = _sanitize_stripped_pe_certificates(staging_root)
    _write_assets(staging_root)
    (staging_root / "AppxManifest.xml").write_text(render_manifest(config), encoding="utf-8", newline="\n")
    return sanitized


def _normalize_timestamps(root: Path, epoch: int) -> None:
    paths = sorted(root.rglob("*"), key=lambda path: len(path.parts), reverse=True)
    for path in paths:
        os.utime(path, (epoch, epoch))
    os.utime(root, (epoch, epoch))


def _safe_reset_work_root(path: Path) -> Path:
    resolved = path.resolve()
    _require(resolved != Path(resolved.anchor), "msix_work_root_is_filesystem_root")
    _require(resolved != Path.home().resolve(), "msix_work_root_is_home")
    _require(len(resolved.parts) >= 3, "msix_work_root_too_broad")
    if resolved.exists():
        _require(resolved.is_dir() and not resolved.is_symlink(), "msix_work_root_invalid")
        shutil.rmtree(resolved)
    resolved.mkdir(parents=True)
    return resolved


def find_makeappx(explicit: Path | None = None) -> Path:
    candidates: list[Path] = []
    if explicit is not None:
        candidates.append(explicit)
    from_environment = os.environ.get("MAKEAPPX_EXE")
    if from_environment:
        candidates.append(Path(from_environment))

    program_files_x86 = os.environ.get("PROGRAMFILES(X86)")
    if program_files_x86:
        sdk_bin = Path(program_files_x86) / "Windows Kits" / "10" / "bin"
        if sdk_bin.is_dir():
            candidates.extend(
                sorted(sdk_bin.glob("10.*.*/x64/makeappx.exe"), key=lambda path: path.parts[-3], reverse=True)
            )
    on_path = shutil.which("makeappx.exe")
    if on_path:
        candidates.append(Path(on_path))

    for candidate in candidates:
        if candidate.is_file():
            return candidate.resolve()
    raise MsixBuildError("makeappx_not_found")


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def package_content_identity(path: Path) -> tuple[int, str]:
    """Hash package paths and uncompressed bytes while ignoring ZIP timestamps."""

    package_digest = hashlib.sha256()
    seen: set[str] = set()
    with zipfile.ZipFile(path) as package:
        entries = package.infolist()
        for entry in entries:
            _require(entry.filename not in seen, "msix_package_duplicate_entry")
            seen.add(entry.filename)
            name_bytes = entry.filename.encode("utf-8")
            package_digest.update(len(name_bytes).to_bytes(4, "big"))
            package_digest.update(name_bytes)
            package_digest.update(entry.file_size.to_bytes(8, "big"))
            entry_digest = hashlib.sha256()
            with package.open(entry, "r") as stream:
                for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                    entry_digest.update(chunk)
            package_digest.update(entry_digest.digest())
    return len(entries), package_digest.hexdigest()


def build_msix(
    *,
    payload_root: Path,
    output: Path,
    work_root: Path,
    config: Mapping[str, Any],
    makeappx: Path,
) -> dict[str, object]:
    safe_work_root = _safe_reset_work_root(work_root)
    staging_root = safe_work_root / "staging"
    sanitized_pe_certificates = prepare_layout(payload_root.resolve(), staging_root, config)
    _normalize_timestamps(staging_root, int(config["reproducibilityEpoch"]))

    output = output.resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    completed = subprocess.run(
        [str(makeappx), "pack", "/d", str(staging_root), "/p", str(output), "/o"],
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if completed.returncode != 0:
        details = (completed.stdout + "\n" + completed.stderr).strip()
        raise MsixBuildError(f"makeappx_failed:{completed.returncode}:{details}")
    _require(output.is_file(), "msix_output_missing")
    entry_count, content_sha256 = package_content_identity(output)

    return {
        "schemaVersion": 1,
        "storeId": config["storeId"],
        "identity": config["identity"],
        "sourceVersion": config["sourceVersion"],
        "packageVersion": config["packageVersion"],
        "assetName": output.name,
        "bytes": output.stat().st_size,
        "sha256": _sha256(output),
        "entryCount": entry_count,
        "contentSha256": content_sha256,
        "sanitizedPeCertificateDirectories": list(sanitized_pe_certificates),
    }


def _write_metadata(path: Path, metadata: Mapping[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(metadata, ensure_ascii=False, indent=2) + "\n", encoding="utf-8", newline="\n")


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--payload-root", type=Path, required=True)
    parser.add_argument("--config", type=Path, required=True)
    output_group = parser.add_mutually_exclusive_group(required=True)
    output_group.add_argument("--output", type=Path)
    output_group.add_argument("--output-root", type=Path)
    parser.add_argument("--work-root", type=Path, required=True)
    parser.add_argument("--source-version", required=True)
    parser.add_argument("--makeappx", type=Path)
    parser.add_argument("--metadata-output", type=Path)
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        config = read_config(args.config.resolve())
        _require(args.source_version == config["sourceVersion"], "msix_source_version_mismatch")
        output = args.output or (args.output_root / config["assetName"])
        makeappx = find_makeappx(args.makeappx)
        metadata = build_msix(
            payload_root=args.payload_root,
            output=output,
            work_root=args.work_root,
            config=config,
            makeappx=makeappx,
        )
        if args.metadata_output is not None:
            _write_metadata(args.metadata_output.resolve(), metadata)
    except MsixBuildError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    print(json.dumps(metadata, ensure_ascii=False, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
