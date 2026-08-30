from __future__ import annotations

import hashlib
import json
import subprocess
import zipfile
from copy import deepcopy
from pathlib import Path

import pytest
from scripts.release import build_production_candidate as production
from scripts.release import v4_package_input_build as v4_build

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]
MANIFEST = ROOT / "release" / "windows-production-manifest.v1.json"


def test_windows_production_manifest_has_one_offline_asset_and_fixed_epoch() -> None:
    manifest = production.read_manifest(MANIFEST)
    assert manifest["product"]["version"] == "0.9.1"
    assert manifest["product"]["releaseTagFormat"] == "{version}"
    assert manifest["product"]["cliJsonProtocolVersion"] == 3
    assert manifest["build"]["mode"] == "pure-python"
    assert manifest["build"]["reproducibilityEpoch"] == 1704067200
    assert manifest["channels"]["offlineZip"]["timestamp"] == "fixed-reproducibility-epoch"
    assert manifest["channels"]["offlineZip"]["assetName"] == "DocWen-windows-x64.zip"
    assert "DocWenCLI-windows-x64.zip" not in json.dumps(manifest)
    assert manifest["releaseAssetAllowlist"] == [
        "DocWen-windows-x64.zip",
        "DocWenCLI-0.9.1-linux-x64.tar.gz",
        "DocWen-0.9.1-linux-x64.tar.gz",
        "SHA256SUMS.txt",
    ]
    assert manifest["toolchain"]["python"] == {
        "implementation": "CPython",
        "fullVersion": "3.12.13",
        "arch": "x86_64",
        "baseExecutableSha256": "02c3bcf63782cc34665ff39ea73d029128ef0849c5a67fe4bbb03748a63fb4f1",
    }
    assert manifest["payload"]["allowlist"] == {
        "path": "release/windows-payload-allowlist.v1.json",
        "status": "FROZEN",
        "sha256": "7c3dbf66160ac6be0c30b9f5f545f2079174fbb5cd34d3783eb7872c7106efae",
        "rejectMissing": True,
        "rejectUnexpected": True,
        "rejectCaseFoldCollision": True,
        "rejectLinksAndNonRegular": True,
    }


def test_source_contracts_and_frozen_payload_allowlist_are_current() -> None:
    manifest = production.read_manifest(MANIFEST)
    production.verify_source_contracts(ROOT, manifest)
    allowlist = ROOT / manifest["payload"]["allowlist"]["path"]
    assert manifest["payload"]["allowlist"]["status"] == "FROZEN"
    assert production.sha256_file(allowlist) == manifest["payload"]["allowlist"]["sha256"]


def test_toolchain_hashes_the_base_interpreter_not_the_path_bound_venv_launcher(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    launcher = tmp_path / "venv-python.exe"
    base = tmp_path / "python.exe"
    uv = tmp_path / "uv.exe"
    launcher.write_bytes(b"path-bound virtual environment launcher")
    base.write_bytes(b"stable managed CPython base executable")
    uv.write_bytes(b"pinned uv executable")
    manifest = {
        "toolchain": {
            "python": {
                "implementation": "CPython",
                "fullVersion": "3.12.13",
                "arch": "x86_64",
                "baseExecutableSha256": production.sha256_file(base),
            },
            "uv": {
                "version": "0.12.0",
                "build": "b88d7c5c4",
                "binarySha256": production.sha256_file(uv),
            },
            "pyinstallerVersion": "6.16.0",
        }
    }

    def fake_run(command: list[str], *, cwd: Path, env: dict[str, str] | None = None) -> str:
        del cwd, env
        if command[0] == str(launcher):
            return json.dumps(
                {
                    "implementation": "CPython",
                    "version": "3.12.13",
                    "arch": "AMD64",
                    "executable": str(launcher),
                    "baseExecutable": str(base),
                }
            )
        if command[0] == str(uv):
            return "uv 0.12.0 (b88d7c5c4 2026-07-28 x86_64-pc-windows-msvc)\n"
        raise AssertionError(command)

    monkeypatch.setattr(production, "run", fake_run)
    monkeypatch.setattr(production.importlib.metadata, "version", lambda name: "6.16.0")

    assert production.verify_toolchain(manifest, launcher, uv) == {
        "pythonBaseExecutable": production.sha256_file(base),
        "uv": production.sha256_file(uv),
    }


def test_build_environment_is_bound_to_the_verified_base_interpreter(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    uv = tmp_path / "uv.exe"
    base = tmp_path / "python.exe"
    clone = tmp_path / "source"
    clone_python = clone / ".venv" / "Scripts" / "python.exe"
    uv.write_bytes(b"pinned uv")
    base.write_bytes(b"verified base python")
    clone_python.parent.mkdir(parents=True)
    clone_python.write_bytes(b"project environment launcher")
    commands: list[list[str]] = []

    def fake_run(command: list[str], *, cwd: Path, env: dict[str, str] | None = None) -> str:
        assert cwd == clone
        assert env == {"UV_OFFLINE": "1"}
        commands.append(command)
        if command[0] == str(clone_python):
            return f"{base}\n"
        return ""

    monkeypatch.setattr(production, "run", fake_run)
    monkeypatch.setattr(Path, "samefile", lambda self, other: self == other)

    assert production.sync_project_environment(uv, clone, base, {"UV_OFFLINE": "1"}) == clone_python
    assert commands[0] == [
        str(uv),
        "sync",
        "--python",
        str(base),
        "--frozen",
        "--offline",
        "--extra",
        "dev",
    ]


def test_windows_build_environment_excludes_incidental_host_tool_paths(tmp_path: Path) -> None:
    clone_python = tmp_path / "clone" / ".venv" / "Scripts" / "python.exe"
    base_python = tmp_path / "python" / "python.exe"
    system_root = tmp_path / "Windows"
    clone_python.parent.mkdir(parents=True)
    base_python.parent.mkdir(parents=True)
    (system_root / "System32").mkdir(parents=True)
    clone_python.write_bytes(b"launcher")
    base_python.write_bytes(b"base")
    env = {
        "SYSTEMROOT": str(system_root),
        "PATH": r"C:\Users\runner\Android\jdk\bin;C:\host-tools",
        "SOURCE_DATE_EPOCH": "1704067200",
    }

    sanitized = production.windows_build_environment(env, clone_python, base_python)

    assert sanitized["PATH"].split(";") == [
        str(clone_python.parent.resolve()),
        str(base_python.parent.resolve()),
        str((system_root / "System32").resolve()),
        str(system_root.resolve()),
    ]
    assert "Android" not in sanitized["PATH"]
    assert sanitized["SOURCE_DATE_EPOCH"] == "1704067200"


def test_v4_toolchain_hashes_the_same_base_interpreter(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    launcher = tmp_path / "venv-python.exe"
    base = tmp_path / "python.exe"
    uv = tmp_path / "uv.exe"
    lock = tmp_path / "uv.lock"
    launcher.write_bytes(b"another path-bound virtual environment launcher")
    base.write_bytes(b"the same stable managed CPython base executable")
    uv.write_bytes(b"the same pinned uv executable")
    lock.write_bytes(b"the exact lock file")
    manifest = {
        "schemaVersion": 1,
        "manifestId": "docwen-windows-production",
        "product": {"version": "0.9.1"},
        "toolchain": {
            "python": {
                "implementation": "CPython",
                "fullVersion": "3.12.13",
                "arch": "x86_64",
                "baseExecutableSha256": production.sha256_file(base),
            },
            "uv": {
                "version": "0.12.0",
                "build": "b88d7c5c4",
                "binarySha256": production.sha256_file(uv),
            },
            "lock": {"path": lock.name, "sha256": production.sha256_file(lock)},
            "pyinstallerVersion": "6.16.0",
        },
    }

    def fake_require_command(command: list[str], *, cwd: Path, label: str) -> subprocess.CompletedProcess[bytes]:
        del cwd
        if label == "python_runtime_probe":
            stdout = json.dumps(
                {
                    "implementation": "CPython",
                    "version": "3.12.13",
                    "arch": "AMD64",
                    "executable": str(launcher),
                    "baseExecutable": str(base),
                }
            ).encode()
        elif label == "uv_version":
            stdout = b"uv 0.12.0 (b88d7c5c4 2026-07-28 x86_64-pc-windows-msvc)\n"
        else:
            raise AssertionError((command, label))
        return subprocess.CompletedProcess(command, 0, stdout=stdout, stderr=b"")

    monkeypatch.setattr(v4_build, "_toolchain_policy", lambda repo: manifest)
    monkeypatch.setattr(v4_build, "require_command", fake_require_command)

    _policy, identity = v4_build._verify_base_toolchain(tmp_path, launcher, uv)
    assert identity["python"]["sha256"] == production.sha256_file(base)


def test_machine_contract_files_pin_exact_d2_closure() -> None:
    manifest = production.read_manifest(MANIFEST)
    files = manifest["sourceContracts"]["machineContractFiles"]
    assert [item["path"] for item in files] == list(production.MACHINE_CONTRACT_PATHS)
    assert len(files) == 32
    assert all(production.sha256_lf_source_file(ROOT / item["path"]) == item["sha256"] for item in files)


def test_machine_contract_files_fail_closed_when_an_expected_trace_is_missing() -> None:
    manifest = deepcopy(production.read_manifest(MANIFEST))
    manifest["sourceContracts"]["machineContractFiles"].pop()
    with pytest.raises(production.ProductionBuildError, match="machine_contract_files_count"):
        production.verify_source_contracts(ROOT, manifest)


def test_machine_contract_files_fail_closed_when_a_hash_drifts() -> None:
    manifest = deepcopy(production.read_manifest(MANIFEST))
    target = manifest["sourceContracts"]["machineContractFiles"][2]
    target["sha256"] = "0" * 64
    with pytest.raises(production.ProductionBuildError, match="source_contract_hash_mismatch"):
        production.verify_source_contracts(ROOT, manifest)


def test_deterministic_zip_normalizes_order_time_and_modes(tmp_path: Path) -> None:
    payload = tmp_path / "payload"
    payload.mkdir()
    (payload / "DocWen.exe").write_bytes(b"gui")
    (payload / "DocWenCLI.exe").write_bytes(b"cli")
    rows = production.capture_payload(payload)
    first = tmp_path / "first.zip"
    second = tmp_path / "second.zip"
    production.deterministic_zip(payload, first, rows, 1_700_000_001)
    production.deterministic_zip(payload, second, rows, 1_700_000_001)
    assert hashlib.sha256(first.read_bytes()).digest() == hashlib.sha256(second.read_bytes()).digest()
    with zipfile.ZipFile(first) as archive:
        assert archive.namelist() == ["DocWen.exe", "DocWenCLI.exe"]
        assert archive.comment == b""
        assert all(info.extra == b"" for info in archive.infolist())


def test_capture_payload_is_canonical_and_requires_both_entry_points(tmp_path: Path) -> None:
    payload = tmp_path / "payload"
    payload.mkdir()
    (payload / "DocWen.exe").write_bytes(b"gui")
    (payload / "DocWenCLI.exe").write_bytes(b"cli")
    rows = production.capture_payload(payload)
    assert [row["path"] for row in rows] == ["DocWen.exe", "DocWenCLI.exe"]
    assert all(row["executable"] is True for row in rows)


@pytest.mark.parametrize(
    "relative,error",
    [
        ("_internal/api-ms-win-core-console-l1-1-0.dll", "payload_host_api_set_forbidden"),
        ("_internal/ucrtbase.dll", "payload_host_ucrt_forbidden"),
    ],
)
def test_capture_payload_rejects_host_selected_windows_runtime(tmp_path: Path, relative: str, error: str) -> None:
    payload = tmp_path / "payload"
    (payload / "_internal").mkdir(parents=True)
    (payload / "DocWen.exe").write_bytes(b"gui")
    (payload / "DocWenCLI.exe").write_bytes(b"cli")
    (payload / relative).write_bytes(b"host-selected runtime")

    with pytest.raises(production.ProductionBuildError, match=error):
        production.capture_payload(payload)


def test_payload_allowlist_mismatch_details_are_bounded_and_actionable() -> None:
    expected = {
        "schemaVersion": 1,
        "entries": [
            {"path": "DocWen.exe", "bytes": 3, "sha256": "a" * 64, "executable": True},
            {"path": "removed.txt", "bytes": 1, "sha256": "b" * 64, "executable": False},
        ],
    }
    observed = {
        "schemaVersion": 1,
        "entries": [
            {"path": "DocWen.exe", "bytes": 4, "sha256": "c" * 64, "executable": True},
            {"path": "added.txt", "bytes": 2, "sha256": "d" * 64, "executable": False},
        ],
    }

    details = production.payload_allowlist_mismatch_details(production.canonical_bytes(expected), observed, limit=1)

    assert details == {
        "addedCount": 1,
        "addedPaths": ["added.txt"],
        "removedCount": 1,
        "removedPaths": ["removed.txt"],
        "changedCount": 1,
        "changes": [
            {
                "path": "DocWen.exe",
                "expected": expected["entries"][0],
                "observed": observed["entries"][0],
            }
        ],
        "truncated": False,
        "observedSha256": hashlib.sha256(production.canonical_bytes(observed)).hexdigest(),
    }


def test_packaged_record_normalization_removes_only_external_rows(tmp_path: Path) -> None:
    payload = tmp_path / "payload"
    record = payload / "_internal" / "example-1.0.dist-info" / "RECORD"
    record.parent.mkdir(parents=True)
    record.write_text(
        "package.py,sha256=stable,1\n../../Scripts/example.exe,sha256=volatile,2\nexample-1.0.dist-info/RECORD,,\n",
        encoding="utf-8",
    )
    assert production.normalize_packaged_record_files(payload) == {
        "changedFiles": 1,
        "removedRows": 1,
    }
    assert record.read_text(encoding="utf-8") == ("package.py,sha256=stable,1\nexample-1.0.dist-info/RECORD,,\n")


def test_packaged_msvc_runtime_normalization_uses_locked_dependency_files(tmp_path: Path) -> None:
    payload = tmp_path / "payload"
    internal = payload / "_internal"
    dependency = tmp_path / "site-packages" / "pikepdf"
    internal.mkdir(parents=True)
    dependency.mkdir(parents=True)
    expected: dict[str, bytes] = {}
    for index, (source_name, target_name) in enumerate(production.PACKAGED_MSVC_RUNTIME_FILES):
        source_bytes = f"locked-runtime-{index}".encode()
        (dependency / source_name).write_bytes(source_bytes)
        (internal / target_name).write_bytes(f"host-runtime-{index}".encode())
        expected[target_name] = source_bytes

    result = production.normalize_packaged_msvc_runtime(payload, dependency)

    assert result["sourcePackage"] == "pikepdf"
    assert [item["path"] for item in result["files"]] == [
        f"_internal/{target_name}" for _source_name, target_name in production.PACKAGED_MSVC_RUNTIME_FILES
    ]
    assert all((internal / name).read_bytes() == contents for name, contents in expected.items())


def test_packaged_msvc_runtime_normalization_fails_before_partial_replacement(tmp_path: Path) -> None:
    payload = tmp_path / "payload"
    internal = payload / "_internal"
    dependency = tmp_path / "site-packages" / "pikepdf"
    internal.mkdir(parents=True)
    dependency.mkdir(parents=True)
    first_target: Path | None = None
    for index, (source_name, target_name) in enumerate(production.PACKAGED_MSVC_RUNTIME_FILES):
        target = internal / target_name
        target.write_bytes(f"host-runtime-{index}".encode())
        if first_target is None:
            first_target = target
        if index != len(production.PACKAGED_MSVC_RUNTIME_FILES) - 1:
            (dependency / source_name).write_bytes(f"locked-runtime-{index}".encode())

    assert first_target is not None
    original = first_target.read_bytes()
    missing_source = production.PACKAGED_MSVC_RUNTIME_FILES[-1][0]
    with pytest.raises(production.ProductionBuildError, match=f"packaged_msvc_runtime_source_missing:{missing_source}"):
        production.normalize_packaged_msvc_runtime(payload, dependency)

    assert first_target.read_bytes() == original


def test_production_work_cleanup_requires_its_exact_lease(tmp_path: Path) -> None:
    work = (tmp_path / "work").resolve()
    work.mkdir()
    production._write_work_lease(work, state="active")
    (work / "intermediate.txt").write_text("temporary\n", encoding="utf-8")

    production._cleanup_owned_work_root(work)

    assert not work.exists()


def test_production_work_cleanup_rejects_foreign_content_root(tmp_path: Path) -> None:
    work = (tmp_path / "work").resolve()
    work.mkdir()
    (work / production.PRODUCTION_WORK_LEASE).write_text(
        '{"schemaVersion":1,"owner":"foreign","root":"untrusted"}\n',
        encoding="utf-8",
    )

    with pytest.raises(production.ProductionBuildError, match="work_root_cleanup_owner_mismatch"):
        production._cleanup_owned_work_root(work)

    assert work.is_dir()
