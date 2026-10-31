"""Compose, ad-hoc sign and verify macOS archives from one existing frozen build."""

from __future__ import annotations

import argparse
import importlib.metadata
import json
import os
import platform
import plistlib
import re
import shutil
import subprocess
import sys
import tarfile
import tomllib
from pathlib import Path
from typing import Any

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from scripts.release.macos_archive import archive_payload, inventory, verify_extracted, write_inventory
from scripts.release.macos_functional import run_cases
from scripts.release.macos_functional_cases import MARKDOWN, TOKEN
from scripts.release.publication_contract import PublicationError, canonical_json, file_identity, require
from scripts.release.publication_download import atomic_record
from tools.run_lease import OwnedRunLease, lease_payload, managed_run

REPO = Path(__file__).resolve().parents[2]
MACHO_MAGICS = {
    b"\xcf\xfa\xed\xfe",
    b"\xfe\xed\xfa\xcf",
    b"\xca\xfe\xba\xbe",
    b"\xbe\xba\xfe\xca",
    b"\xca\xfe\xba\xbf",
    b"\xbf\xba\xfe\xca",
}


def read_contract(path: Path) -> dict[str, Any]:
    contract = json.loads(path.read_text(encoding="utf-8"))
    version = tomllib.loads((REPO / "pyproject.toml").read_text(encoding="utf-8"))["project"]["version"]
    require(
        contract.get("schemaVersion") == 1 and contract.get("manifestId") == "docwen-macos-production",
        "invalid macOS contract",
    )
    require(contract.get("product") == {"name": "DocWen", "version": version}, "macOS contract version mismatch")
    require(
        contract.get("target")
        == {"os": "macos", "minimumVersion": "14.0", "arch": "arm64", "platformTag": "macos-arm64"},
        "unsupported macOS target",
    )
    require(
        contract.get("signing") == {"identity": "-", "developerId": False, "notarized": False},
        "ad-hoc signing required",
    )
    require(
        contract.get("bundle")
        == {"name": "DocWen.app", "identifier": "io.github.ZHYX91.DocWen", "payload": "Contents/Resources/payload"},
        "unexpected app layout",
    )
    require(
        contract.get("archive") == {"format": "tar.gz", "epoch": 1704067200, "symlinks": "relative-internal-only"},
        "unexpected archive policy",
    )
    require(
        contract.get("artifacts")
        == [
            {"id": kind, "name": f"{prefix}-{version}-macos-arm64.tar.gz", "root": f"{prefix}-{version}-macos-arm64"}
            for kind, prefix in (("gui-cli", "DocWen"), ("cli-only", "DocWenCLI"))
        ],
        "unexpected macOS assets",
    )
    return contract


def command(arguments: list[str], *, environment: dict[str, str] | None = None, timeout: int = 120) -> str:
    result = subprocess.run(arguments, env=environment, capture_output=True, text=True, timeout=timeout)
    require(result.returncode == 0, f"command failed: {arguments[0]}: {result.stderr[-4000:]}")
    return result.stdout + result.stderr


def macho_files(root: Path) -> list[Path]:
    files = []
    for relative, record in inventory(root).items():
        if record["type"] == "file":
            with (root / relative).open("rb") as stream:
                if stream.read(4) in MACHO_MAGICS:
                    files.append(root / relative)
    return files


def framework_subject(path: Path, root: Path) -> Path:
    for parent in path.parents:
        if parent == root:
            break
        if parent.name.endswith(".framework"):
            return parent
    return path


def sign_payload(root: Path) -> None:
    """Seal frameworks inside out after COLLECT; signing a bare dylib is insufficient."""
    frameworks: set[Path] = set()
    for path in macho_files(root):
        for parent in path.parents:
            if parent == root:
                break
            if parent.name.endswith(".framework"):
                frameworks.add(parent)
        if framework_subject(path, root) == path:
            command(["/usr/bin/codesign", "--force", "--sign", "-", "--timestamp=none", str(path)])
    for framework in sorted(frameworks, key=lambda path: (-len(path.parts), str(path))):
        command(["/usr/bin/codesign", "--force", "--sign", "-", "--timestamp=none", str(framework)])


def minimum_macos_version(load_commands: str, relative: str) -> tuple[int, int, int]:
    """Read the arm64 deployment floor, independently of the build SDK version."""
    versions: list[tuple[int, int, int]] = []
    for block in re.split(r"(?m)^\s*Load command \d+\s*$", load_commands):
        commands = re.findall(r"(?m)^\s*cmd\s+(\S+)\s*$", block)
        relevant = [name for name in commands if name == "LC_BUILD_VERSION" or name.startswith("LC_VERSION_MIN_")]
        if not relevant:
            continue
        require(len(commands) == len(relevant) == 1, f"ambiguous deployment metadata: {relative}")
        name = relevant[0]
        if name == "LC_BUILD_VERSION":
            platforms = re.findall(r"(?m)^\s*platform\s+(\S+)\s*$", block)
            require(platforms in (["MACOS"], ["1"]), f"non-macOS or missing platform metadata: {relative}")
            field = "minos"
        else:
            require(name == "LC_VERSION_MIN_MACOSX", f"non-macOS deployment metadata: {relative}")
            field = "version"
        values = re.findall(rf"(?m)^\s*{field}\s+(\S+)\s*$", block)
        require(
            len(values) == 1 and re.fullmatch(r"\d+(?:\.\d+){0,2}", values[0]) is not None,
            f"missing or ambiguous deployment version: {relative}",
        )
        parts = [int(part) for part in values[0].split(".")]
        parts += [0] * (3 - len(parts))
        versions.append((parts[0], parts[1], parts[2]))
    require(len(versions) == 1, f"missing or ambiguous deployment metadata: {relative}")
    version = versions[0]
    require(
        (0, 0, 0) < version <= (14, 0, 0),
        f"macOS minimum {'.'.join(map(str, version))} exceeds 14.0 or is invalid: {relative}",
    )
    return version


def verify_signatures(root: Path, *, app: Path | None = None) -> dict[str, Any]:
    binaries = []
    minimum_versions = []
    verified: set[Path] = set()
    for path in macho_files(root):
        relative = path.relative_to(root).as_posix()
        architectures = command(["/usr/bin/lipo", "-archs", str(path)]).strip().split()
        require("arm64" in architectures, f"non-arm64 code: {relative}")
        subject = framework_subject(path, root)
        if subject not in verified:
            command(["/usr/bin/codesign", "--verify", "--deep", "--strict", str(subject)])
            verified.add(subject)
        try:
            load_commands = command(["/usr/bin/otool", "-arch", "arm64", "-l", str(path)])
        except PublicationError as error:
            raise PublicationError(f"cannot read arm64 deployment metadata: {relative}: {error}") from error
        minimum_versions.append(minimum_macos_version(load_commands, relative))
        binaries.append(relative)
    require(bool(binaries), "no Mach-O code in package")
    if app is not None:
        command(["/usr/bin/codesign", "--verify", "--deep", "--strict", str(app)])
        description = command(["/usr/bin/codesign", "--display", "--verbose=4", str(app)])
        require(
            "Signature=adhoc" in description and "TeamIdentifier=not set" in description,
            "unexpected app signing identity",
        )
    for name in ("DocWen", "DocWenCLI"):
        for relative in binaries:
            if Path(relative).name == name:
                require(
                    "Signature=adhoc" in command(["/usr/bin/codesign", "-d", "--verbose=4", str(root / relative)]),
                    "entry point must be ad-hoc signed",
                )
    maximum = max(minimum_versions)
    version_text = ".".join(map(str, maximum if maximum[2] else maximum[:2]))
    return {
        "machOCount": len(binaries),
        "architecture": "arm64",
        "appSignature": "adhoc" if app else None,
        "maximumMachOMinimumVersion": version_text,
    }


def compose_app(payload: Path, root: Path, version: str) -> Path:
    app = root / "DocWen.app"
    contents = app / "Contents"
    macos, resources = contents / "MacOS", contents / "Resources"
    macos.mkdir(parents=True)
    resources.mkdir()
    inventory(payload)  # Validate links before copying anything.
    shutil.copytree(payload, resources / "payload", symlinks=True)
    shutil.copyfile(REPO / "docs/user-guides/macos.md", root / "FIRST-OPEN.md")
    (contents / "Info.plist").write_bytes(
        plistlib.dumps(
            {
                "CFBundleDevelopmentRegion": "en",
                "CFBundleExecutable": "DocWen",
                "CFBundleName": "DocWen",
                "CFBundleIdentifier": "io.github.ZHYX91.DocWen",
                "CFBundlePackageType": "APPL",
                "CFBundleShortVersionString": version,
                "CFBundleVersion": version,
                "LSMinimumSystemVersion": "14.0",
                "NSHighResolutionCapable": True,
                "CFBundleIconFile": "DocWen.icns",
            },
            sort_keys=True,
        )
    )
    from PIL import Image

    with Image.open(REPO / "assets/icon.png") as image:
        image.save(resources / "DocWen.icns", format="ICNS")
    command(
        [
            "/usr/bin/clang",
            "-arch",
            "arm64",
            "-mmacosx-version-min=14.0",
            "-O2",
            str(REPO / "scripts/release/macos_launcher.c"),
            "-o",
            str(macos / "DocWen"),
        ]
    )
    sign_payload(resources / "payload")
    command(["/usr/bin/codesign", "--force", "--sign", "-", "--timestamp=none", str(app)])
    return app


def package(dist: Path, output: Path, report: Path) -> None:
    require(platform.system() == "Darwin" and platform.machine() == "arm64", "native Apple Silicon required")
    contract = read_contract(REPO / "release/macos-production-manifest.v1.json")
    version = contract["product"]["version"]
    require(not output.exists() and not report.exists(), "package outputs must be new")
    (REPO / "build").mkdir(exist_ok=True)
    with managed_run(
        REPO / "build", prefix="macos-release-", owner="docwen.release.macos", kind="macos-package"
    ) as run:
        results: dict[str, Any] = {
            "schema": "docwen.macos.release-checks.v1",
            "host": {"system": platform.system(), "release": platform.mac_ver()[0], "machine": platform.machine()},
            "software": {
                "python": platform.python_version(),
                "uv": command(["uv", "--version"]).strip(),
                "pyinstaller": importlib.metadata.version("pyinstaller"),
                "libreoffice": command(["/Applications/LibreOffice.app/Contents/MacOS/soffice", "--version"]).strip(),
            },
            "sourceCommit": command(["git", "-C", str(REPO), "rev-parse", "HEAD"]).strip(),
            "contract": file_identity(REPO / "release/macos-production-manifest.v1.json"),
            "lockfile": file_identity(REPO / "uv.lock"),
            "packages": {},
            "limitations": [
                "No browser quarantine or manual Gatekeeper exception acceptance",
                "No interactive menu, drag/drop or physical-device acceptance",
                "No Developer ID signing or notarization",
            ],
        }
        archives = []
        for artifact in contract["artifacts"]:
            root = run.root / artifact["root"]
            root.mkdir()
            app = None
            if artifact["id"] == "gui-cli":
                app = compose_app(dist / f"DocWen_v{version}_macos-arm64", root, version)
            else:
                payload = dist / f"DocWenCLI_v{version}_macos-arm64"
                inventory(payload)
                shutil.copytree(payload, root, dirs_exist_ok=True, symlinks=True)
                shutil.copyfile(REPO / "docs/user-guides/macos.md", root / "FIRST-OPEN.md")
                sign_payload(root)
            signatures = verify_signatures(root, app=app)
            write_inventory(root)
            archive = run.root / artifact["name"]
            archive_identity = archive_payload(root, archive, top=artifact["root"], epoch=contract["archive"]["epoch"])
            extracted = run.root / (artifact["id"] + "-extracted")
            extracted.mkdir()
            with tarfile.open(archive) as bundle:
                bundle.extractall(extracted, filter="data")
            installed = extracted / artifact["root"]
            verify_extracted(installed)
            extracted_app = installed / "DocWen.app" if app else None
            require(verify_signatures(installed, app=extracted_app) == signatures, "signature drift after extraction")
            binary_dir = installed / "DocWen.app/Contents/Resources/payload" if app else installed
            cases = run_cases(binary_dir / "DocWenCLI", run.root / (artifact["id"] + "-cases"), REPO)
            if app:
                # Test the actual bundle launcher too, independent of the caller's cwd.
                environment = os.environ.copy()
                environment.update(QT_QPA_PLATFORM="offscreen", DOCWEN_GUI_DISABLE_STATE_SAVE="1")
                # Keep verifier fixtures and subprocess scratch paths under this
                # owned, canonical root rather than macOS's /var temp aliases.
                gui_temp = run.root / "gui-temp"
                gui_temp.mkdir()
                environment.update({key: str(gui_temp) for key in ("TMPDIR", "TMP", "TEMP")})
                command(
                    [
                        sys.executable,
                        str(REPO / "scripts/release/verify_packaged_gui.py"),
                        "--binary-dir",
                        str(binary_dir),
                        "--settings-smoke",
                    ],
                    environment=environment,
                    timeout=180,
                )
                environment.update(
                    DOCWEN_DATA_DIR=str(run.root / "launch-profile"),
                    DOCWEN_RUNTIME_ROOT=str(run.root / "launch-runtime"),
                    DOCWEN_GUI_TEST_AUTOCLOSE_MS="2000",
                )
                command([str(installed / "DocWen.app/Contents/MacOS/DocWen")], environment=environment, timeout=45)
                # LaunchServices + the native Cocoa plugin, using a fresh profile.
                # This is local CI launch evidence, not browser quarantine approval.
                native_report = run.root / "native-settings.json"
                default_home = run.root / "default-home"
                default_home.mkdir()
                default_environment = os.environ.copy()
                for key in (
                    "DOCWEN_DATA_DIR",
                    "DOCWEN_CONFIG_DIR",
                    "DOCWEN_LOG_DIR",
                    "DOCWEN_LOG_TO_TEMP",
                    "PYTHONPATH",
                    "PYTHONHOME",
                    "XDG_DATA_HOME",
                    "XDG_CONFIG_HOME",
                    "XDG_CACHE_HOME",
                    "XDG_STATE_HOME",
                    "XDG_RUNTIME_DIR",
                ):
                    default_environment.pop(key, None)
                default_environment.update(
                    HOME=str(default_home), DOCWEN_RUNTIME_ROOT=str(run.root / "default-runtime")
                )
                default_input = run.root / "default-profile.md"
                default_input.write_text(MARKDOWN, encoding="utf-8")
                # User configs are sparse and deliberately not backfilled on read.
                # Observe a seeded override instead of assuming startup creates configs/.
                profile = default_home / "Library/Application Support/docwen"
                (profile / "configs").mkdir(parents=True)
                profile_config = profile / "configs/logger.toml"
                profile_config.write_text(
                    'enable = true\nfile_prefix = "mac-profile-proof"\ndirectory_mode = "user"\n', encoding="utf-8"
                )
                profile_config_identity = file_identity(profile_config)
                converted = subprocess.run(
                    [
                        str(binary_dir / "DocWenCLI"),
                        "convert",
                        str(default_input),
                        "--to",
                        "docx",
                        "--output-dir",
                        str(run.root / "default-output"),
                        "--json",
                    ],
                    env=default_environment,
                    capture_output=True,
                    text=True,
                    timeout=90,
                )
                require(converted.returncode == 0, "default-profile conversion failed: " + converted.stderr[-2000:])
                default_payload = json.loads(converted.stdout)
                from docx import Document

                default_document = Document(default_payload["data"]["output"])
                require(
                    any(TOKEN in paragraph.text for paragraph in default_document.paragraphs),
                    "default-profile output content missing",
                )
                require(
                    (profile / "logs/mac-profile-proof.log").is_file(),
                    "default-profile logger override was not consumed",
                )
                native_environment = {
                    "QT_QPA_PLATFORM": "cocoa",
                    "DOCWEN_GUI_DISABLE_STATE_SAVE": "1",
                    "DOCWEN_GUI_TEST_AUTOCLOSE_MS": "5000",
                    "DOCWEN_GUI_TEST_SETTINGS_REPORT": str(native_report),
                    "HOME": str(default_home),
                    "DOCWEN_DATA_DIR": "",
                    "DOCWEN_CONFIG_DIR": "",
                    "DOCWEN_LOG_DIR": "",
                    "DOCWEN_LOG_TO_TEMP": "",
                    "XDG_DATA_HOME": "",
                    "XDG_CONFIG_HOME": "",
                    "XDG_CACHE_HOME": "",
                    "XDG_STATE_HOME": "",
                    "XDG_RUNTIME_DIR": "",
                    "DOCWEN_RUNTIME_ROOT": str(run.root / "native-runtime"),
                }
                launch = ["/usr/bin/open", "-W", "-n"]
                for key, value in native_environment.items():
                    launch.extend(["--env", f"{key}={value}"])
                command([*launch, str(extracted_app)], timeout=90)
                from scripts.release.verify_packaged_gui import _verify_settings_smoke_report

                _verify_settings_smoke_report(native_report)
                require(file_identity(profile_config) == profile_config_identity, "sparse user config was modified")
                environment["QT_QPA_PLATFORM"] = "cocoa"
                command(
                    [
                        sys.executable,
                        str(REPO / "scripts/release/verify_packaged_gui.py"),
                        "--binary-dir",
                        str(binary_dir),
                        "--gongwen-fields-smoke",
                    ],
                    environment=environment,
                    timeout=300,
                )
            verify_extracted(installed)
            verify_signatures(installed, app=extracted_app)
            require(file_identity(archive) == archive_identity, "candidate archive changed during tests")
            results["packages"][artifact["id"]] = {
                "asset": artifact["name"],
                **archive_identity,
                "signatures": signatures,
                "cases": cases,
                "guiSettingsAndLauncher": "passed-offscreen-and-native-launchservices-cocoa"
                if app
                else "not-applicable",
                "nativeGuiConversion": "six-gongwen-fields-cases-passed" if app else "not-applicable",
                "defaultUserProfile": "passed-with-isolated-home" if app else "not-applicable",
            }
            archives.append(archive)
        _publish_archives(archives, output=output, report=report, results=results)


def _publish_archives(archives: list[Path], *, output: Path, report: Path, results: dict[str, Any]) -> None:
    """Publish one already verified build without leaving anonymous partial archives."""
    require(not output.exists() and not output.is_symlink(), "package outputs must be new")
    require(not report.exists() and not report.is_symlink(), "package report must be new")
    require(len({archive.name for archive in archives}) == len(archives), "archive names must be unique")
    require(report not in {output / archive.name for archive in archives}, "package report conflicts with archive")
    output.parent.mkdir(parents=True, exist_ok=True)
    identities = {archive.name: file_identity(archive) for archive in archives}
    with managed_run(
        output.parent, prefix=".docwen-macos-publish-", owner="docwen.release.macos", kind="macos-publication"
    ) as publication:
        atomic_record(
            publication.root / "publication.json",
            canonical_json(
                {
                    "output": str(output),
                    "archives": identities,
                    "report": {"path": str(report), "content": results},
                }
            ),
        )
        for archive in archives:
            staged = publication.root / archive.name
            shutil.copyfile(archive, staged)
            require(file_identity(staged) == identities[archive.name], "candidate archive copy mismatch")
        output.mkdir()
        root_identity = output.lstat()
        try:
            atomic_record(
                publication.root / "candidate-identity.json",
                canonical_json(
                    {
                        "root": str(output),
                        "rootIdentity": {"device": root_identity.st_dev, "inode": root_identity.st_ino},
                        "owner": "docwen.release.macos",
                        "kind": "macos-candidate",
                    }
                ),
            )
            lease = lease_payload(output, owner="docwen.release.macos", kind="macos-candidate")
            lease["publicationRoot"] = str(publication.root)
            lease["artifacts"] = identities
            lease["report"] = str(report)
            marker = output / ".docwen-temp-lease.json"
            atomic_record(marker, canonical_json(lease))
            marker_identity = marker.lstat()
        except BaseException as error:
            error.add_note(
                f"Candidate reservation: {output}; device={root_identity.st_dev}; inode={root_identity.st_ino}; "
                f"publication evidence: {publication.root}"
            )
            try:
                current = output.lstat()
                if (current.st_dev, current.st_ino) == (
                    root_identity.st_dev,
                    root_identity.st_ino,
                ) and not output.is_symlink():
                    # Only an empty directory can be retired here. A partial
                    # marker or unexpected entry remains with its reservation.
                    output.rmdir()
            except OSError as cleanup_error:
                error.add_note(f"Candidate initialization cleanup retained: {cleanup_error}")
            raise

        candidate = OwnedRunLease(
            output, root_identity, marker_identity, lease, "docwen.release.macos", canonical_json(lease)
        )

        try:
            for archive in archives:
                os.link(publication.root / archive.name, output / archive.name)
            report.parent.mkdir(parents=True, exist_ok=True)
            atomic_record(report, canonical_json(results))
            for archive in archives:
                require(file_identity(output / archive.name) == identities[archive.name], "published archive changed")
        except BaseException as error:
            outcome = "retained-interrupted" if isinstance(error, KeyboardInterrupt) else "retained-failure"
            try:
                candidate.save(outcome)
            except (OSError, ValueError, PublicationError) as recording_error:
                error.add_note(f"Could not record candidate failure; inspect {publication.root}: {recording_error}")
                candidate.record_recovery(
                    error,
                    "failure-recording",
                    outcome=outcome,
                    recording_error=recording_error,
                    execution_error=error,
                )
            raise
        try:
            candidate.save("completed-success")
        except (OSError, ValueError, PublicationError) as error:
            # All archive bytes and the receipt are committed. A terminal lease
            # recording failure must not rewrite that fact as a failed build.
            candidate.record_recovery(
                error,
                "terminal-recording",
                outcome="completed-success",
                recording_error=error,
            )
            raise


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--dist", type=Path, default=REPO / "dist")
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--report", type=Path, required=True)
    args = parser.parse_args()
    package(args.dist.resolve(), args.output.resolve(), args.report.resolve())


if __name__ == "__main__":
    main()
