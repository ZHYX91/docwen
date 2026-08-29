#!/usr/bin/env python3
"""Build the single manifest-bound Windows production payload.

Normal mode fails closed until the checked-in payload allowlist is calibrated.
Calibration mode creates an observed allowlist as evidence but never calls the
result production-ready.  The source checkout is never used as a build/output
directory: a clean local clone is created under the caller-owned work root.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.metadata
import json
import os
import re
import shutil
import stat
import subprocess
import sys
import tomllib
import zipfile
from collections.abc import Iterable, Sequence
from datetime import UTC, datetime
from pathlib import Path, PurePosixPath
from typing import Any

if __package__ in {None, ""}:
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from scripts.build.payload_normalization import normalize_packaged_record_files

EMPTY_SHA256 = hashlib.sha256(b"").hexdigest()
PRODUCTION_WORK_LEASE = ".docwen-temp-lease.json"
STABLE_SEMVER = re.compile(r"^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$")
SHA256_HEX = re.compile(r"^[0-9a-f]{64}$")
MACHINE_CONTRACT_PATHS = (
    "contracts/schemas/docwen.machine.v1.schema.json",
    "contracts/schemas/docwen.machine.diagnostic_evidence.v1.schema.json",
    "contracts/schemas/docwen.artifact_bundle.v2.schema.json",
    "contracts/schemas/docwen.proofread_report.v2.schema.json",
    "contracts/schemas/docwen.semantic_bibliography.v1.schema.json",
    "contracts/schemas/docwen.round_trip_sidecar.v1.schema.json",
    "contracts/conformance-manifest.json",
    "contracts/fixtures/valid/semantic-bibliography.rich.json",
    "contracts/fixtures/invalid/semantic-bibliography.unknown-field.json",
    "contracts/fixtures/valid/machine.capability-list.response.json",
    "contracts/fixtures/valid/machine.task-lifecycle.trace.json",
    "contracts/fixtures/valid/machine.v4-diagnostic-evidence.trace.json",
    "contracts/fixtures/valid/artifact-bundle.ocr.json",
    "contracts/fixtures/invalid/artifact-bundle.missing-page-fragment-semantics.json",
    "contracts/fixtures/invalid/artifact-bundle.unexpected-page-semantics.json",
    "contracts/fixtures/invalid/artifact-bundle.invalid-page-range.json",
    "contracts/fixtures/invalid/artifact-bundle.page-ordinal-mismatch.json",
    "contracts/fixtures/invalid/artifact-bundle.page-count-mismatch.json",
    "contracts/fixtures/invalid/artifact-bundle.duplicate-page-index.json",
    "contracts/fixtures/invalid/artifact-bundle.incomplete-page-sequence.json",
    "contracts/fixtures/invalid/artifact-bundle.page-source-mismatch.json",
    "contracts/fixtures/invalid/artifact-bundle.resource-page-mismatch.json",
    "contracts/fixtures/invalid/artifact-bundle.page-semantics-unknown-field.json",
    "contracts/fixtures/invalid/artifact-bundle.duplicate-relation-ordinal.json",
    "contracts/fixtures/invalid/machine.dangling-diagnostic-artifact.trace.json",
    "contracts/fixtures/invalid/machine.partial-diagnostic-evidence.json",
    "contracts/fixtures/invalid/machine.duplicate-input-logical-path.trace.json",
    "contracts/fixtures/invalid/machine.input-slot-cardinality-mismatch.trace.json",
    "contracts/fixtures/invalid/machine.input-slot-kind-mismatch.trace.json",
    "contracts/fixtures/invalid/machine.input-slot-media-type-mismatch.trace.json",
    "contracts/fixtures/invalid/machine.invalid-input-logical-path.trace.json",
    "contracts/fixtures/invalid/machine.undeclared-input-role.trace.json",
)
PACKAGED_MSVC_RUNTIME_FILES = (
    ("msvcp140_1.dll", "MSVCP140_1.dll"),
    ("msvcp140.dll", "msvcp140.dll"),
    ("vcruntime140.dll", "vcruntime140.dll"),
    ("vcruntime140_1.dll", "vcruntime140_1.dll"),
)


class ProductionBuildError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ProductionBuildError(message)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def sha256_lf_source_file(path: Path) -> str:
    """Hash a governed text source using the repository's LF contract."""

    payload = path.read_bytes().replace(b"\r\n", b"\n")
    require(b"\r" not in payload, f"source_contract_bare_carriage_return:{path}")
    return hashlib.sha256(payload).hexdigest()


def canonical_bytes(value: object) -> bytes:
    return (json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n").encode("utf-8")


def atomic_write(path: Path, payload: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.tmp")
    with temporary.open("xb") as stream:
        stream.write(payload)
        stream.flush()
        os.fsync(stream.fileno())
    os.replace(temporary, path)


def _write_work_lease(work: Path, *, state: str) -> None:
    atomic_write(
        work / PRODUCTION_WORK_LEASE,
        canonical_bytes(
            {
                "schemaVersion": 1,
                "owner": "docwen.release.build-production-candidate",
                "kind": "production-build-work",
                "pid": os.getpid(),
                "createdAt": datetime.now(UTC).isoformat().replace("+00:00", "Z"),
                "state": state,
                "root": str(work),
            }
        ),
    )


def _update_work_lease(work: Path, *, state: str) -> None:
    marker = work / PRODUCTION_WORK_LEASE
    payload = json.loads(marker.read_text(encoding="utf-8"))
    payload["state"] = state
    marker.unlink()
    atomic_write(marker, canonical_bytes(payload))


def _cleanup_owned_work_root(work: Path) -> None:
    safe_work = work.resolve(strict=True)
    metadata = safe_work.lstat()
    reparse_flag = int(getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0))
    require(
        not stat.S_ISLNK(metadata.st_mode)
        and not (reparse_flag and int(getattr(metadata, "st_file_attributes", 0)) & reparse_flag),
        "work_root_cleanup_reparse_rejected",
    )
    marker = safe_work / PRODUCTION_WORK_LEASE
    require(marker.is_file() and not marker.is_symlink(), "work_root_cleanup_lease_missing")
    payload = json.loads(marker.read_text(encoding="utf-8"))
    require(payload.get("owner") == "docwen.release.build-production-candidate", "work_root_cleanup_owner_mismatch")
    require(payload.get("root") == str(safe_work), "work_root_cleanup_identity_mismatch")
    shutil.rmtree(safe_work)
    require(not safe_work.exists(), "work_root_cleanup_failed")


def run(command: Sequence[str], *, cwd: Path, env: dict[str, str] | None = None) -> str:
    completed = subprocess.run(
        command,
        cwd=cwd,
        env=env,
        text=True,
        encoding="utf-8",
        errors="replace",
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        check=False,
    )
    if completed.returncode != 0:
        tail = "\n".join(completed.stdout.splitlines()[-80:])
        raise ProductionBuildError(f"command_failed:{completed.returncode}:{command[0]}\n{tail}")
    return completed.stdout


def git(repo: Path, *args: str) -> str:
    return run(["git", *args], cwd=repo).strip()


def read_manifest(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ProductionBuildError("production_manifest_unreadable") from exc
    require(isinstance(value, dict), "production_manifest_not_object")
    require(value.get("schemaVersion") == 1, "production_manifest_schema_version")
    require(value.get("manifestId") == "docwen-windows-production", "production_manifest_id")
    return value


def verify_machine_contract_files(repo: Path, contracts: dict[str, Any]) -> dict[str, str]:
    """Verify the D2 machine contract closure pinned by the release policy.

    The list is intentionally exact and ordered.  A production build must not
    silently omit a negative trace or accept a different machine contract set.
    """

    items = contracts.get("machineContractFiles")
    if not isinstance(items, list):
        raise ProductionBuildError("machine_contract_files_missing")
    require(len(items) == len(MACHINE_CONTRACT_PATHS), "machine_contract_files_count")
    hashes: dict[str, str] = {}
    for expected_path, item in zip(MACHINE_CONTRACT_PATHS, items, strict=True):
        require(isinstance(item, dict), f"machine_contract_file_invalid:{expected_path}")
        require(item.get("path") == expected_path, f"machine_contract_file_path_drift:{expected_path}")
        expected_hash = item.get("sha256")
        require(
            isinstance(expected_hash, str) and SHA256_HEX.fullmatch(expected_hash) is not None,
            f"machine_contract_file_hash_invalid:{expected_path}",
        )
        path = repo / expected_path
        require(path.is_file(), f"machine_contract_file_missing:{expected_path}")
        actual = sha256_lf_source_file(path)
        require(actual == expected_hash, f"source_contract_hash_mismatch:{expected_path}")
        hashes[expected_path] = actual
    return hashes


def verify_source_contracts(repo: Path, manifest: dict[str, Any]) -> dict[str, str]:
    product = manifest["product"]
    version = product["version"]
    require(
        isinstance(version, str) and STABLE_SEMVER.fullmatch(version) is not None, "product_version_not_stable_semver"
    )
    require(product["releaseTagFormat"] == "{version}", "release_tag_must_not_use_v_prefix")
    require(manifest["build"]["mode"] == "pure-python", "production_build_mode_must_be_pure_python")
    require(manifest["build"]["reproducibilityEpoch"] == 1704067200, "reproducibility_epoch_drift")
    require(
        manifest["target"] == {"os": "windows", "arch": "x86_64", "platformTag": "win-x64"}, "production_target_drift"
    )
    for relative in manifest["sourceContracts"]["versionProjectionPaths"]:
        data = tomllib.loads((repo / relative).read_text(encoding="utf-8"))
        require(data["project"]["version"] == version, f"version_projection_mismatch:{relative}")
    hashes: dict[str, str] = {}
    contracts = manifest["sourceContracts"]
    for item in [contracts["protocol"], contracts["schema"], *contracts["fixtures"], manifest["toolchain"]["lock"]]:
        path = repo / item["path"]
        actual = sha256_lf_source_file(path)
        require(actual == item["sha256"], f"source_contract_hash_mismatch:{item['path']}")
        hashes[item["path"]] = actual
    hashes.update(verify_machine_contract_files(repo, contracts))
    protocol_source = (repo / contracts["protocol"]["path"]).read_text(encoding="utf-8")
    match = re.search(r"^PROTOCOL_VERSION\s*=\s*(\d+)\s*$", protocol_source, re.MULTILINE)
    require(
        match is not None and int(match.group(1)) == product["cliJsonProtocolVersion"],
        "cli_json_protocol_projection_mismatch",
    )
    expected_assets = [
        "DocWen-windows-x64.zip",
        f"DocWenCLI-{version}-linux-x64.tar.gz",
        f"DocWen-{version}-linux-x64.tar.gz",
        "SHA256SUMS.txt",
    ]
    require(manifest["releaseAssetAllowlist"] == expected_assets, "release_asset_allowlist_drift")
    require(manifest["channels"]["offlineZip"]["assetName"] == expected_assets[0], "offline_asset_name_drift")
    return hashes


def _verify_toolchain(manifest: dict[str, Any], python: Path, uv: Path) -> tuple[dict[str, str], Path]:
    expected_python = manifest["toolchain"]["python"]
    probe_source = (
        "import json, platform, sys; "
        "print(json.dumps({"
        "'implementation': platform.python_implementation(), "
        "'version': platform.python_version(), "
        "'arch': platform.machine(), "
        "'executable': sys.executable, "
        "'baseExecutable': sys._base_executable"
        "}, sort_keys=True))"
    )
    try:
        python_probe = json.loads(run([str(python), "-I", "-c", probe_source], cwd=python.parent))
        require(
            isinstance(python_probe, dict)
            and set(python_probe) == {"arch", "baseExecutable", "executable", "implementation", "version"}
            and all(isinstance(value, str) and value for value in python_probe.values()),
            "python_runtime_probe_invalid",
        )
        reported_python = Path(python_probe["executable"]).resolve(strict=True)
        base_python = Path(python_probe["baseExecutable"]).resolve(strict=True)
    except (json.JSONDecodeError, OSError, TypeError) as exc:
        raise ProductionBuildError("python_runtime_probe_invalid") from exc
    require(python.samefile(reported_python), "python_executable_identity_mismatch")
    base_metadata = base_python.lstat()
    reparse_flag = int(getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0))
    require(
        stat.S_ISREG(base_metadata.st_mode)
        and not base_python.is_symlink()
        and not (reparse_flag and int(getattr(base_metadata, "st_file_attributes", 0)) & reparse_flag),
        "python_base_executable_not_regular",
    )
    require(python_probe["implementation"] == expected_python["implementation"], "python_implementation_mismatch")
    require(python_probe["version"] == expected_python["fullVersion"], "python_version_mismatch")
    python_hash = sha256_file(base_python)
    require(python_hash == expected_python["baseExecutableSha256"], "python_base_executable_hash_mismatch")
    require(python_probe["arch"].lower() in {"amd64", "x86_64"}, "python_arch_mismatch")
    uv_text = run([str(uv), "--version"], cwd=Path.cwd()).strip()
    expected_uv = manifest["toolchain"]["uv"]
    require(f"uv {expected_uv['version']} ({expected_uv['build']} " in uv_text, "uv_version_mismatch")
    require(sha256_file(uv) == expected_uv["binarySha256"], "uv_binary_hash_mismatch")
    require(
        importlib.metadata.version("pyinstaller") == manifest["toolchain"]["pyinstallerVersion"],
        "pyinstaller_version_mismatch",
    )
    return {"pythonBaseExecutable": python_hash, "uv": sha256_file(uv)}, base_python


def verify_toolchain(manifest: dict[str, Any], python: Path, uv: Path) -> dict[str, str]:
    hashes, _base_python = _verify_toolchain(manifest, python, uv)
    return hashes


def sync_project_environment(uv: Path, clone: Path, base_python: Path, env: dict[str, str]) -> Path:
    """Create the build environment from the exact verified base interpreter."""

    run(
        [
            str(uv),
            "sync",
            "--python",
            str(base_python),
            "--frozen",
            "--offline",
            "--extra",
            "dev",
        ],
        cwd=clone,
        env=env,
    )
    clone_python = clone / ".venv" / "Scripts" / "python.exe"
    require(clone_python.is_file(), "clone_python_missing")
    clone_base_text = run(
        [str(clone_python), "-I", "-c", "import sys; print(sys._base_executable)"],
        cwd=clone,
        env=env,
    ).strip()
    try:
        clone_base = Path(clone_base_text).resolve(strict=True)
    except OSError as exc:
        raise ProductionBuildError("clone_python_base_probe_invalid") from exc
    require(base_python.samefile(clone_base), "clone_python_base_identity_mismatch")
    return clone_python


def windows_build_environment(env: dict[str, str], clone_python: Path, base_python: Path) -> dict[str, str]:
    """Remove incidental developer and runner tools from PyInstaller's DLL search path."""

    try:
        system_root_value = env.get("SYSTEMROOT") or env["SystemRoot"]
        system_root = Path(system_root_value).resolve(strict=True)
        system32 = (system_root / "System32").resolve(strict=True)
        clone_scripts = clone_python.parent.resolve(strict=True)
        base_root = base_python.parent.resolve(strict=True)
    except (KeyError, OSError) as exc:
        raise ProductionBuildError("windows_build_path_invalid") from exc
    sanitized = dict(env)
    sanitized["PATH"] = ";".join(str(path) for path in (clone_scripts, base_root, system32, system_root))
    return sanitized


def capture_payload(payload: Path) -> list[dict[str, object]]:
    require(payload.is_dir(), "payload_missing")
    rows: list[dict[str, object]] = []
    folded: dict[str, str] = {}
    for path in sorted(payload.rglob("*"), key=lambda item: item.relative_to(payload).as_posix().encode("utf-8")):
        relative = path.relative_to(payload).as_posix()
        require(not path.is_symlink(), f"payload_link_forbidden:{relative}")
        require(path.is_file() or path.is_dir(), f"payload_nonregular_forbidden:{relative}")
        folded_relative = relative.casefold()
        require(
            not (folded_relative.startswith("_internal/api-ms-win-") and folded_relative.endswith(".dll")),
            f"payload_host_api_set_forbidden:{relative}",
        )
        require(folded_relative != "_internal/ucrtbase.dll", f"payload_host_ucrt_forbidden:{relative}")
        key = relative.casefold()
        require(
            key not in folded or folded[key] == relative, f"payload_casefold_collision:{folded.get(key)}:{relative}"
        )
        folded[key] = relative
        if path.is_file():
            rows.append(
                {
                    "path": relative,
                    "bytes": path.stat().st_size,
                    "sha256": sha256_file(path),
                    "executable": relative in {"DocWen.exe", "DocWenCLI.exe"},
                }
            )
    require(
        {row["path"] for row in rows}.issuperset({"DocWen.exe", "DocWenCLI.exe"}), "payload_root_executables_missing"
    )
    require(
        not any(str(row["path"]).startswith("assets/screenshots/") for row in rows), "payload_screenshots_forbidden"
    )
    return rows


def payload_allowlist_mismatch_details(
    expected_bytes: bytes,
    observed_allowlist: dict[str, object],
    *,
    limit: int = 32,
) -> dict[str, object]:
    """Return a bounded, log-safe summary for an exact allowlist mismatch."""

    require(limit > 0, "payload_allowlist_diff_limit_invalid")

    def index_entries(value: object, *, label: str) -> dict[str, dict[str, object]]:
        if not isinstance(value, list):
            raise ProductionBuildError(f"payload_allowlist_{label}_entries_invalid")
        indexed: dict[str, dict[str, object]] = {}
        for entry in value:
            if not isinstance(entry, dict):
                raise ProductionBuildError(f"payload_allowlist_{label}_entry_invalid")
            path = entry.get("path")
            if not isinstance(path, str) or not path:
                raise ProductionBuildError(f"payload_allowlist_{label}_path_invalid")
            indexed[path] = entry
        return indexed

    try:
        expected_allowlist = json.loads(expected_bytes)
        if not isinstance(expected_allowlist, dict):
            raise ProductionBuildError("payload_allowlist_expected_invalid")
        expected_by_path = index_entries(expected_allowlist.get("entries"), label="expected")
        observed_by_path = index_entries(observed_allowlist.get("entries"), label="observed")
    except (json.JSONDecodeError, KeyError, TypeError) as exc:
        raise ProductionBuildError("payload_allowlist_diff_unavailable") from exc

    added_paths = sorted(observed_by_path.keys() - expected_by_path.keys())
    removed_paths = sorted(expected_by_path.keys() - observed_by_path.keys())
    changed_paths = sorted(
        path
        for path in expected_by_path.keys() & observed_by_path.keys()
        if expected_by_path[path] != observed_by_path[path]
    )
    changes = [
        {
            "path": path,
            "expected": expected_by_path[path],
            "observed": observed_by_path[path],
        }
        for path in changed_paths[:limit]
    ]
    return {
        "addedCount": len(added_paths),
        "addedPaths": added_paths[:limit],
        "removedCount": len(removed_paths),
        "removedPaths": removed_paths[:limit],
        "changedCount": len(changed_paths),
        "changes": changes,
        "truncated": any(len(paths) > limit for paths in (added_paths, removed_paths, changed_paths)),
        "observedSha256": hashlib.sha256(canonical_bytes(observed_allowlist)).hexdigest(),
    }


def normalize_packaged_msvc_runtime(payload: Path, dependency_root: Path) -> dict[str, object]:
    """Replace host-selected MSVC runtime files with locked wheel copies.

    PyInstaller resolves these four DLLs through the Windows loader search
    path.  A host Visual C++ Runtime update can therefore change an otherwise
    manifest-bound payload.  The locked pikepdf wheel ships the same runtime
    closure, so use those files as the deterministic source and let the frozen
    payload allowlist verify their exact hashes.
    """

    internal = payload / "_internal"
    pairs: list[tuple[Path, Path, str]] = []
    for source_name, target_name in PACKAGED_MSVC_RUNTIME_FILES:
        source = dependency_root / source_name
        target = internal / target_name
        require(source.is_file() and not source.is_symlink(), f"packaged_msvc_runtime_source_missing:{source_name}")
        require(target.is_file() and not target.is_symlink(), f"packaged_msvc_runtime_target_missing:{target_name}")
        pairs.append((source, target, target_name))

    files: list[dict[str, object]] = []
    for source, target, target_name in pairs:
        shutil.copyfile(source, target)
        files.append(
            {
                "path": f"_internal/{target_name}",
                "bytes": target.stat().st_size,
                "sha256": sha256_file(target),
            }
        )
    return {"sourcePackage": "pikepdf", "files": files}


def deterministic_zip(payload: Path, destination: Path, rows: Iterable[dict[str, object]], epoch: int) -> None:
    from datetime import datetime

    timestamp = datetime.fromtimestamp(max(epoch, 315532800), tz=UTC)
    date_time = (
        timestamp.year,
        timestamp.month,
        timestamp.day,
        timestamp.hour,
        timestamp.minute,
        timestamp.second - timestamp.second % 2,
    )
    with zipfile.ZipFile(
        destination, "x", compression=zipfile.ZIP_DEFLATED, compresslevel=9, strict_timestamps=True
    ) as archive:
        archive.comment = b""
        for row in rows:
            relative = str(row["path"])
            info = zipfile.ZipInfo(relative, date_time=date_time)
            info.create_system = 3
            mode = 0o755 if bool(row["executable"]) else 0o644
            info.external_attr = (stat.S_IFREG | mode) << 16
            info.compress_type = zipfile.ZIP_DEFLATED
            info.flag_bits |= 0x800
            archive.writestr(
                info,
                (payload / PurePosixPath(relative)).read_bytes(),
                compress_type=zipfile.ZIP_DEFLATED,
                compresslevel=9,
            )


def build(args: argparse.Namespace) -> dict[str, Any]:
    repo = args.repo.resolve(strict=True)
    manifest_path = args.manifest.resolve(strict=True)
    output = args.output_root.resolve(strict=False)
    work = args.work_root.resolve(strict=False)
    require(not output.exists(), "output_root_already_exists")
    require(not work.exists(), "work_root_already_exists")
    require(repo not in output.parents and repo != output, "output_root_inside_source")
    require(repo not in work.parents and repo != work, "work_root_inside_source")
    manifest = read_manifest(manifest_path)
    require(git(repo, "status", "--porcelain=v2") == "", "source_must_be_clean")
    commit = git(repo, "rev-parse", "HEAD")
    tree = git(repo, "rev-parse", "HEAD^{tree}")
    epoch = int(git(repo, "show", "-s", "--format=%ct", "HEAD"))
    reproducibility_epoch = int(manifest["build"]["reproducibilityEpoch"])
    source_hashes = verify_source_contracts(repo, manifest)
    tool_hashes, base_python = _verify_toolchain(
        manifest,
        args.python.resolve(strict=True),
        args.uv.resolve(strict=True),
    )
    output.mkdir(parents=True)
    work.mkdir(parents=True)
    _write_work_lease(work, state="active")
    clone = work / "source"
    run(["git", "clone", "--local", "--no-hardlinks", "--no-checkout", str(repo), str(clone)], cwd=work)
    run(["git", "checkout", "--detach", commit], cwd=clone)
    require(
        git(clone, "rev-parse", "HEAD") == commit and git(clone, "rev-parse", "HEAD^{tree}") == tree,
        "clone_identity_mismatch",
    )
    env = dict(os.environ)
    env.update(
        {
            "PYTHONHASHSEED": "0",
            "TZ": "UTC",
            "DOCWEN_GUI_DISABLE_STATE_SAVE": "1",
            "SOURCE_DATE_EPOCH": str(reproducibility_epoch),
            "UV_OFFLINE": "1",
        }
    )
    clone_python = sync_project_environment(args.uv, clone, base_python, env)
    build_env = windows_build_environment(env, clone_python, base_python)
    build_log = run(
        [str(clone_python), "scripts/build/build.py", "--skip-cython", "--version", manifest["product"]["version"]],
        cwd=clone,
        env=build_env,
    )
    evidence = output / "evidence"
    evidence.mkdir()
    atomic_write(evidence / "primitive-build.log", build_log.encode("utf-8"))
    deploy = clone / "dist" / f"DocWen_v{manifest['product']['version']}_win-x64"
    require(deploy.is_dir(), "unified_deploy_directory_missing")
    payload = output / "payload"
    shutil.copytree(deploy, payload, ignore=shutil.ignore_patterns("screenshots"))
    for readme in sorted(clone.glob(manifest["payload"]["readmeSourceGlob"])):
        shutil.copy2(readme, payload / readme.name)
    msvc_runtime_normalization = normalize_packaged_msvc_runtime(
        payload,
        clone / ".venv" / "Lib" / "site-packages" / "pikepdf",
    )
    record_normalization = normalize_packaged_record_files(payload)
    rows = capture_payload(payload)
    allowlist_policy = manifest["payload"]["allowlist"]
    observed_allowlist = {"schemaVersion": 1, "entries": rows}
    observed_bytes = canonical_bytes(observed_allowlist)
    if args.calibrate_allowlist:
        require(
            allowlist_policy["status"] == "CALIBRATION_REQUIRED" and allowlist_policy["sha256"] is None,
            "calibration_not_authorized_by_manifest",
        )
        atomic_write(evidence / "windows-payload-allowlist.v1.json", observed_bytes)
        classification = "CALIBRATION_ONLY_NOT_PRODUCTION"
    else:
        require(
            allowlist_policy["status"] == "FROZEN" and isinstance(allowlist_policy["sha256"], str),
            "payload_allowlist_not_frozen",
        )
        allowlist_path = clone / allowlist_policy["path"]
        require(sha256_file(allowlist_path) == allowlist_policy["sha256"], "payload_allowlist_file_hash_mismatch")
        expected_allowlist_bytes = allowlist_path.read_bytes()
        if expected_allowlist_bytes != observed_bytes:
            details = payload_allowlist_mismatch_details(expected_allowlist_bytes, observed_allowlist)
            raise ProductionBuildError(
                f"payload_allowlist_mismatch:{json.dumps(details, ensure_ascii=True, sort_keys=True, separators=(',', ':'))}"
            )
        classification = "PRODUCTION_PAYLOAD_VERIFIED"
    zip_path = output / manifest["channels"]["offlineZip"]["assetName"]
    deterministic_zip(payload, zip_path, rows, reproducibility_epoch)
    payload_manifest = {"schemaVersion": 1, "files": rows}
    payload_manifest_bytes = canonical_bytes(payload_manifest)
    atomic_write(evidence / "payload-manifest.json", payload_manifest_bytes)
    resolved = {
        "schemaVersion": 1,
        "classification": classification,
        "source": {
            "commit": commit,
            "tree": tree,
            "commitEpoch": epoch,
            "reproducibilityEpoch": reproducibility_epoch,
            "emptyStatusSha256": EMPTY_SHA256,
        },
        "policyManifest": {
            "path": manifest_path.relative_to(repo).as_posix(),
            "sha256": sha256_file(manifest_path),
            "gitBlob": git(repo, "hash-object", str(manifest_path)),
        },
        "sourceContracts": source_hashes,
        "toolchain": tool_hashes,
        "payloadManifestSha256": hashlib.sha256(payload_manifest_bytes).hexdigest(),
        "payloadAllowlistSha256": hashlib.sha256(observed_bytes).hexdigest(),
        "packagedMsvcRuntimeNormalization": msvc_runtime_normalization,
        "packagedRecordNormalization": record_normalization,
        "offlineZip": {"name": zip_path.name, "bytes": zip_path.stat().st_size, "sha256": sha256_file(zip_path)},
    }
    resolved_bytes = canonical_bytes(resolved)
    atomic_write(evidence / "resolved-build-inputs.json", resolved_bytes)
    require(git(repo, "status", "--porcelain=v2") == "", "source_changed_during_build")
    if getattr(args, "keep_work_root", False):
        _update_work_lease(work, state="retained-manual")
    else:
        _cleanup_owned_work_root(work)
    print(json.dumps(resolved, ensure_ascii=False, sort_keys=True, separators=(",", ":")))
    return resolved


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(description=__doc__)
    result.add_argument("--repo", type=Path, required=True)
    result.add_argument("--manifest", type=Path, required=True)
    result.add_argument("--python", type=Path, required=True)
    result.add_argument("--uv", type=Path, required=True)
    result.add_argument("--output-root", type=Path, required=True)
    result.add_argument("--work-root", type=Path, required=True)
    result.add_argument(
        "--keep-work-root",
        action="store_true",
        help="Keep the owned production build work root after a successful build.",
    )
    result.add_argument("--calibrate-allowlist", action="store_true")
    return result


def main(argv: Sequence[str] | None = None) -> int:
    try:
        build(parser().parse_args(argv))
        return 0
    except (ProductionBuildError, OSError, KeyError, TypeError, ValueError) as exc:
        print(f"production build failed closed: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
