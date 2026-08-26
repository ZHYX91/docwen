from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any, cast

from scripts.release import v4_candidate_contract as candidate_contract
from scripts.release import v4_evidence_contract as evidence_contract
from scripts.release import v4_package_input_contract as input_contract

PACKAGE_NAMES = input_contract.PACKAGE_NAMES
BuildOutput = input_contract.BuildOutput
V4PackageInputError = input_contract.V4PackageInputError
_read_json = input_contract._read_json
_identity = input_contract._identity
require_command = input_contract.require_command
git = input_contract.git
sha256_bytes = input_contract.sha256_bytes


def _toolchain_policy(repo: Path) -> dict[str, Any]:
    value, _ = _read_json(
        repo / "release/windows-production-manifest.v1.json",
        root=repo,
        label="windows_production_manifest",
    )
    if value.get("schemaVersion") != 1 or value.get("manifestId") != "docwen-windows-production":
        raise V4PackageInputError("windows_production_manifest_identity_invalid")
    product = value.get("product")
    toolchain = value.get("toolchain")
    if not isinstance(product, dict) or product.get("version") != candidate_contract.PRODUCT_VERSION:
        raise V4PackageInputError("windows_production_product_version_invalid")
    if not isinstance(toolchain, dict):
        raise V4PackageInputError("windows_production_toolchain_missing")
    return value


def _verify_base_toolchain(repo: Path, python: Path, uv: Path) -> tuple[dict[str, Any], dict[str, object]]:
    policy = _toolchain_policy(repo)
    toolchain = cast(dict[str, Any], policy["toolchain"])
    python_policy = cast(dict[str, Any], toolchain["python"])
    uv_policy = cast(dict[str, Any], toolchain["uv"])
    safe_python = evidence_contract.safe_regular_file(python, label="python")
    safe_uv = evidence_contract.safe_regular_file(uv, label="uv")
    uv_identity = _identity(safe_uv, root=safe_uv.parent)
    probe_source = (
        "import json, platform, sys; "
        "print(json.dumps({"
        "'arch': platform.machine(), "
        "'baseExecutable': sys._base_executable, "
        "'executable': sys.executable, "
        "'implementation': platform.python_implementation(), "
        "'version': platform.python_version()"
        "}, sort_keys=True))"
    )
    probe = require_command([str(safe_python), "-I", "-c", probe_source], cwd=repo, label="python_runtime_probe")
    try:
        python_probe = json.loads(probe.stdout.decode("utf-8", errors="strict"))
        if (
            not isinstance(python_probe, dict)
            or set(python_probe) != {"arch", "baseExecutable", "executable", "implementation", "version"}
            or not all(isinstance(value, str) and value for value in python_probe.values())
        ):
            raise ValueError("python runtime probe shape")
        reported_executable = evidence_contract.safe_regular_file(
            Path(python_probe["executable"]), label="python_reported_executable"
        )
        base_python = evidence_contract.safe_regular_file(
            Path(python_probe["baseExecutable"]), label="python_base_executable"
        )
    except (UnicodeDecodeError, json.JSONDecodeError, OSError, TypeError, ValueError) as exc:
        raise V4PackageInputError("python_runtime_probe_invalid") from exc
    if not safe_python.samefile(reported_executable):
        raise V4PackageInputError("python_executable_identity_mismatch")
    python_identity = _identity(base_python, root=base_python.parent)
    reported_python = f"Python {python_probe['version']}"
    uv_version = require_command([str(safe_uv), "--version"], cwd=repo, label="uv_version")
    reported_uv = uv_version.stdout.decode("utf-8", errors="replace").strip()
    if (
        python_probe["implementation"] != python_policy.get("implementation")
        or python_probe["arch"].lower() not in {"amd64", "x86_64"}
        or reported_python != f"Python {python_policy.get('fullVersion')}"
        or python_identity["sha256"] != python_policy.get("baseExecutableSha256")
    ):
        raise V4PackageInputError("python_toolchain_identity_mismatch")
    if not reported_uv.startswith(f"uv {uv_policy.get('version')} ({uv_policy.get('build')} ") or uv_identity[
        "sha256"
    ] != uv_policy.get("binarySha256"):
        raise V4PackageInputError("uv_toolchain_identity_mismatch")
    lock = cast(dict[str, object], toolchain["lock"])
    lock_path = repo / str(lock.get("path"))
    if _identity(lock_path, root=repo)["sha256"] != lock.get("sha256"):
        raise V4PackageInputError("uv_lock_identity_mismatch")
    return policy, {
        "python": {**python_identity, "version": reported_python},
        "uv": {**uv_identity, "version": reported_uv},
        "pyinstallerVersion": toolchain.get("pyinstallerVersion"),
    }


def default_package_builder(clone: Path, python: Path, uv: Path) -> BuildOutput:
    policy, toolchain = _verify_base_toolchain(clone, python, uv)
    environment = dict(os.environ)
    environment.update(
        {
            "PYTHONHASHSEED": "0",
            "TZ": "UTC",
            "DOCWEN_GUI_DISABLE_STATE_SAVE": "1",
            "SOURCE_DATE_EPOCH": str(cast(dict[str, Any], policy["build"])["reproducibilityEpoch"]),
            "UV_OFFLINE": "1",
            "UV_PYTHON": str(python.resolve(strict=True)),
        }
    )
    require_command(
        [str(uv), "sync", "--frozen", "--offline", "--extra", "dev"],
        cwd=clone,
        env=environment,
        label="offline_sync",
        timeout=1800,
    )
    clone_python = clone / ".venv/Scripts/python.exe"
    if not clone_python.is_file():
        raise V4PackageInputError("clone_python_missing")
    pyinstaller = (
        require_command(
            [
                str(clone_python),
                "-c",
                "import importlib.metadata as m; print(m.version('pyinstaller'))",
            ],
            cwd=clone,
            env=environment,
            label="clone_pyinstaller_version",
        )
        .stdout.decode("utf-8", errors="strict")
        .strip()
    )
    if pyinstaller != toolchain["pyinstallerVersion"]:
        raise V4PackageInputError("clone_pyinstaller_version_mismatch")
    build_command = [
        str(clone_python),
        "scripts/build/build.py",
        "--skip-cython",
        "--version",
        candidate_contract.PRODUCT_VERSION,
    ]
    build = require_command(
        build_command,
        cwd=clone,
        env=environment,
        label="primitive_build",
        timeout=3600,
    )
    gui = clone / "dist" / PACKAGE_NAMES[0]
    cli = clone / "dist" / PACKAGE_NAMES[1]
    if not gui.is_dir() or not cli.is_dir():
        raise V4PackageInputError("primitive_build_package_missing")
    if git(clone, "status", "--porcelain=v2", label="post_build_status"):
        raise V4PackageInputError("tracked_source_changed_during_build")
    return BuildOutput(
        gui=gui,
        cli=cli,
        metadata={
            "command": ["<clone-python>", *build_command[1:]],
            "stdoutSha256": sha256_bytes(build.stdout),
            "stdoutBytes": len(build.stdout),
            "toolchain": toolchain,
        },
    )


__all__ = ["default_package_builder"]
