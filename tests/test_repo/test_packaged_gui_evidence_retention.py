"""Safety and success contracts for optional packaged-GUI evidence retention."""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import subprocess
from pathlib import Path

import pytest
from tests.support.release_packaging import use_compact_pymupdf_layout_manifest
from tests.support.release_packaging import (
    write_packaged_common_resources as _write_packaged_common_resources,
)
from tests.support.release_packaging import (
    write_packaged_gui_assets as _write_packaged_gui_assets,
)

pytestmark = pytest.mark.unit


@pytest.fixture(autouse=True)
def _use_compact_pymupdf_layout_manifest(monkeypatch: pytest.MonkeyPatch) -> None:
    use_compact_pymupdf_layout_manifest(monkeypatch)


def _packaged_gui(tmp_path: Path) -> tuple[Path, str]:
    binary_dir = tmp_path / "dist"
    binary_dir.mkdir()
    binary_name = "DocWen.exe"
    (binary_dir / binary_name).write_bytes(b"placeholder")
    _write_packaged_common_resources(binary_dir)
    _write_packaged_gui_assets(binary_dir)
    return binary_dir, binary_name


def _write_office_evidence_fixture(
    verification_dir: Path,
) -> tuple[dict[str, object], Path]:
    from scripts.release import verify_packaged_gui

    office_dir = verification_dir / "gui_office_smoke"
    office_dir.mkdir(parents=True, exist_ok=True)
    cases: list[dict[str, object]] = []
    for case_name, backend in (
        ("docx", "msoffice_word"),
        ("xlsx", "msoffice_excel"),
        ("markdown", "msoffice_word"),
    ):
        case_dir = office_dir / case_name
        case_dir.mkdir()
        case: dict[str, object] = {
            "case": case_name,
            "surface": "action",
            "backend": backend,
            "backendIdentity": {"backend": backend, "version": "16.0.1.2"},
            "checks": {"contentPassed": True, "pageCount": 1, "layoutReview": "not_performed"},
        }
        for subject, content in (
            ("source", f"{case_name}-source".encode()),
            ("output", f"{case_name}-output".encode()),
            ("report", f"{case_name}-report".encode()),
        ):
            file_path = case_dir / f"{subject}.bin"
            file_path.write_bytes(content)
            case[subject] = {
                "path": file_path.relative_to(office_dir).as_posix(),
                "bytes": len(content),
                "sha256": hashlib.sha256(content).hexdigest(),
            }
        cases.append(case)
    evidence: dict[str, object] = {
        "schema": verify_packaged_gui._OFFICE_EVIDENCE_SCHEMA,
        "host": {
            "system": "Windows",
            "release": "11",
            "machine": "AMD64",
            "pythonPlatform": "win32",
        },
        "cases": cases,
    }
    evidence_path = office_dir / "office-smoke-evidence.json"
    verify_packaged_gui._atomic_json_write(evidence_path, evidence)
    return evidence, evidence_path


def _make_directory_link(link: Path, target: Path) -> None:
    if os.name == "nt":
        completed = subprocess.run(
            ["cmd.exe", "/d", "/c", "mklink", "/J", str(link), str(target)],
            capture_output=True,
            text=True,
            errors="replace",
            check=False,
        )
        if completed.returncode != 0:
            raise RuntimeError(f"junction fixture failed: {completed.stdout}\n{completed.stderr}")
        return
    link.symlink_to(target, target_is_directory=True)


def test_evidence_dir_is_parsed_and_preserves_a_verified_success_tree(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_gui

    binary_dir, binary_name = _packaged_gui(tmp_path)
    verification_dir = tmp_path / "isolated-run"
    verification_dir.mkdir()
    evidence_dir = tmp_path / "retained-evidence"
    run_directories: list[Path] = []
    monkeypatch.setattr(
        verify_packaged_gui.tempfile,
        "mkdtemp",
        lambda **_kwargs: str(verification_dir),
    )
    monkeypatch.setattr(verify_packaged_gui, "_optional_workspace_root", lambda _explicit: None)

    def fake_run(binary_path: Path, *args: str, cwd: Path) -> subprocess.CompletedProcess[str]:
        run_directories.append(cwd)
        (cwd / "empty-directory").mkdir()
        (cwd / "artifact.bin").write_bytes(b"\x00DocWen evidence\xff")
        (cwd / "log_home" / "logs").mkdir(parents=True, exist_ok=True)
        (cwd / "log_home" / "logs" / "docwen.log").write_text("ok", encoding="utf-8")
        return subprocess.CompletedProcess([str(binary_path), *args], 0, stdout="", stderr="")

    monkeypatch.setattr(verify_packaged_gui, "_run", fake_run)

    exit_code = verify_packaged_gui.main(
        [
            "--binary-dir",
            str(binary_dir),
            "--binary-name",
            binary_name,
            "--evidence-dir",
            str(evidence_dir),
        ]
    )

    assert exit_code == 0
    assert run_directories == [verification_dir]
    assert not verification_dir.exists()
    assert (evidence_dir / "empty-directory").is_dir()
    artifact = evidence_dir / "artifact.bin"
    assert artifact.read_bytes() == b"\x00DocWen evidence\xff"
    assert hashlib.sha256(artifact.read_bytes()).hexdigest() == (
        "542950c816762d4fda71451a8a6f04ea7691d9139fb5cd69e432680f795865e1"
    )


def test_default_success_still_cleans_without_retaining_evidence(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_gui

    binary_dir, binary_name = _packaged_gui(tmp_path)
    verification_dir = tmp_path / "isolated-run"
    verification_dir.mkdir()
    monkeypatch.setattr(
        verify_packaged_gui.tempfile,
        "mkdtemp",
        lambda **_kwargs: str(verification_dir),
    )
    monkeypatch.setattr(verify_packaged_gui, "_optional_workspace_root", lambda _explicit: None)

    def fake_run(binary_path: Path, *args: str, cwd: Path) -> subprocess.CompletedProcess[str]:
        (cwd / "log_home" / "logs").mkdir(parents=True, exist_ok=True)
        (cwd / "log_home" / "logs" / "docwen.log").write_text("ok", encoding="utf-8")
        return subprocess.CompletedProcess([str(binary_path), *args], 0, stdout="", stderr="")

    monkeypatch.setattr(verify_packaged_gui, "_run", fake_run)

    assert verify_packaged_gui.main(["--binary-dir", str(binary_dir), "--binary-name", binary_name]) == 0
    assert not verification_dir.exists()


def test_evidence_destination_must_be_new_absolute_and_outside_binary_dir(tmp_path: Path) -> None:
    from scripts.release import verify_packaged_gui

    binary_dir, _binary_name = _packaged_gui(tmp_path)
    existing = tmp_path / "existing"
    existing.mkdir()
    sentinel = existing / "sentinel.txt"
    sentinel.write_text("keep", encoding="utf-8")

    with pytest.raises(ValueError, match="evidence_dir_not_absolute"):
        verify_packaged_gui._validate_evidence_destination(Path("relative-evidence"), binary_dir=binary_dir)
    with pytest.raises(FileExistsError, match="evidence_dir_exists"):
        verify_packaged_gui._validate_evidence_destination(existing, binary_dir=binary_dir)
    with pytest.raises(RuntimeError, match="inside_binary_dir"):
        verify_packaged_gui._validate_evidence_destination(binary_dir / "evidence", binary_dir=binary_dir)

    assert sentinel.read_text(encoding="utf-8") == "keep"


def test_evidence_destination_rejects_dotdot_alias_and_linked_parent(tmp_path: Path) -> None:
    from scripts.release import verify_packaged_gui

    binary_dir, _binary_name = _packaged_gui(tmp_path)
    safe_parent = tmp_path / "safe-parent"
    safe_parent.mkdir()
    child = safe_parent / "child"
    child.mkdir()
    alias = child / ".." / "evidence"
    with pytest.raises(RuntimeError, match="evidence_dir_alias_rejected"):
        verify_packaged_gui._validate_evidence_destination(alias, binary_dir=binary_dir)

    real_parent = tmp_path / "real-parent"
    real_parent.mkdir()
    linked_parent = tmp_path / "linked-parent"
    _make_directory_link(linked_parent, real_parent)
    with pytest.raises(RuntimeError, match="evidence_parent_link_rejected"):
        verify_packaged_gui._validate_evidence_destination(linked_parent / "evidence", binary_dir=binary_dir)


def test_evidence_copy_rejects_linked_source_without_following_it(tmp_path: Path) -> None:
    from scripts.release import verify_packaged_gui

    binary_dir, _binary_name = _packaged_gui(tmp_path)
    verification_dir = tmp_path / "verification"
    verification_dir.mkdir()
    external = tmp_path / "external"
    external.mkdir()
    external_file = external / "outside.txt"
    external_file.write_text("outside", encoding="utf-8")
    _make_directory_link(verification_dir / "linked", external)
    evidence_dir = tmp_path / "evidence"

    with pytest.raises(RuntimeError, match="evidence_link_rejected"):
        verify_packaged_gui._preserve_success_evidence(
            verification_dir,
            evidence_dir,
            binary_dir=binary_dir,
        )

    assert verification_dir.exists()
    assert external_file.read_text(encoding="utf-8") == "outside"
    assert not evidence_dir.exists()


def test_evidence_copy_mismatch_fails_closed_and_retains_source(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_gui

    binary_dir, _binary_name = _packaged_gui(tmp_path)
    verification_dir = tmp_path / "verification"
    verification_dir.mkdir()
    source_file = verification_dir / "artifact.bin"
    source_file.write_bytes(b"authoritative")
    evidence_dir = tmp_path / "evidence"
    real_copytree = shutil.copytree

    def corrupting_copytree(source: Path, destination: Path, **kwargs: object) -> Path:
        copied = real_copytree(source, destination, **kwargs)
        (destination / "artifact.bin").write_bytes(b"corrupted")
        return copied

    monkeypatch.setattr(verify_packaged_gui.shutil, "copytree", corrupting_copytree)

    with pytest.raises(RuntimeError, match="evidence_manifest_mismatch"):
        verify_packaged_gui._preserve_success_evidence(
            verification_dir,
            evidence_dir,
            binary_dir=binary_dir,
        )

    assert verification_dir.exists()
    assert source_file.read_bytes() == b"authoritative"
    assert evidence_dir.exists()


def test_main_retains_temporary_source_when_evidence_copy_fails(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    from scripts.release import verify_packaged_gui

    binary_dir, binary_name = _packaged_gui(tmp_path)
    verification_dir = tmp_path / "isolated-run"
    verification_dir.mkdir()
    evidence_dir = tmp_path / "evidence"
    monkeypatch.setattr(
        verify_packaged_gui.tempfile,
        "mkdtemp",
        lambda **_kwargs: str(verification_dir),
    )
    monkeypatch.setattr(verify_packaged_gui, "_optional_workspace_root", lambda _explicit: None)

    def fake_run(binary_path: Path, *args: str, cwd: Path) -> subprocess.CompletedProcess[str]:
        (cwd / "log_home" / "logs").mkdir(parents=True, exist_ok=True)
        (cwd / "log_home" / "logs" / "docwen.log").write_text("ok", encoding="utf-8")
        return subprocess.CompletedProcess([str(binary_path), *args], 0, stdout="", stderr="")

    monkeypatch.setattr(verify_packaged_gui, "_run", fake_run)
    monkeypatch.setattr(
        verify_packaged_gui,
        "_preserve_success_evidence",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(RuntimeError("copy failed")),
    )

    with pytest.raises(RuntimeError, match="copy failed"):
        verify_packaged_gui.main(
            [
                "--binary-dir",
                str(binary_dir),
                "--binary-name",
                binary_name,
                "--evidence-dir",
                str(evidence_dir),
            ]
        )

    assert verification_dir.exists()
    captured = capsys.readouterr()
    assert "packaged_gui_smoke_ok" not in captured.out
    assert f"packaged_gui_failure_artifacts_retained: {verification_dir}" in captured.err
    lease = json.loads((verification_dir / verify_packaged_gui._VERIFICATION_LEASE).read_text(encoding="utf-8"))
    assert lease["state"] == "retained-failure"


def test_governed_success_writes_compact_receipt_and_removes_raw_run(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_gui
    from tools import workspace_root

    binary_dir, binary_name = _packaged_gui(tmp_path)
    governed = tmp_path / ".workspace"
    governed.mkdir()
    (governed / "README.md").write_text("# DocWen 本地工作区\n", encoding="utf-8")
    for name in workspace_root._GOVERNANCE_DIRECTORIES:
        (governed / name).mkdir()
    receipt = governed / "acceptance" / "candidate-packaged-gui.json"
    run_directories: list[Path] = []

    def fake_run(binary_path: Path, *args: str, cwd: Path) -> subprocess.CompletedProcess[str]:
        run_directories.append(cwd)
        (cwd / "log_home" / "logs").mkdir(parents=True, exist_ok=True)
        (cwd / "log_home" / "logs" / "docwen.log").write_text("ok", encoding="utf-8")
        return subprocess.CompletedProcess([str(binary_path), *args], 0, stdout="", stderr="")

    monkeypatch.setattr(verify_packaged_gui, "_run", fake_run)

    assert (
        verify_packaged_gui.main(
            [
                "--binary-dir",
                str(binary_dir),
                "--binary-name",
                binary_name,
                "--workspace-root",
                str(governed),
                "--candidate-id",
                "candidate-1",
                "--receipt-output",
                str(receipt),
            ]
        )
        == 0
    )
    assert len(run_directories) == 1
    assert not run_directories[0].exists()
    payload = json.loads(receipt.read_text(encoding="utf-8"))
    assert payload["schema"] == verify_packaged_gui._RECEIPT_SCHEMA
    assert payload["candidateId"] == "candidate-1"
    assert payload["result"] == "passed"
    assert payload["selectedGates"] == ["default-smoke"]


def test_office_gate_receipt_binds_host_backends_and_evidence_hash(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_gui

    binary_dir, binary_name = _packaged_gui(tmp_path)
    verification_dir = tmp_path / "verification"
    verification_dir.mkdir()
    (verification_dir / "log_home" / "logs").mkdir(parents=True)
    (verification_dir / "log_home" / "logs" / "docwen.log").write_text("ok", encoding="utf-8")
    evidence, evidence_path = _write_office_evidence_fixture(verification_dir)

    payload = verify_packaged_gui._build_acceptance_receipt(
        verification_dir,
        candidate_id="candidate-office",
        binary_path=binary_dir / binary_name,
        selected_gates=["office"],
    )

    office = payload["gateEvidence"]["office"]
    assert office["path"] == "gui_office_smoke/office-smoke-evidence.json"
    assert office["bytes"] == evidence_path.stat().st_size
    assert office["sha256"] == hashlib.sha256(evidence_path.read_bytes()).hexdigest()
    assert office["cases"] == evidence["cases"]
    assert office["host"] == evidence["host"]
    assert office["backends"] == ["msoffice_word", "msoffice_excel", "msoffice_word"]
    assert payload["selectedGates"] == ["office"]


@pytest.mark.parametrize(
    ("mutation", "error_code"),
    [
        ("wrong-hash", "office_evidence_file_mismatch"),
        ("missing-file", "office_evidence_file_missing"),
        ("boolean-bytes", "office_evidence_identity_invalid"),
        ("boolean-page-count", "office_evidence_checks_invalid"),
        ("parent-traversal", "office_evidence_identity_invalid"),
    ],
)
def test_office_gate_receipt_rejects_claims_that_do_not_match_evidence_tree(
    tmp_path: Path,
    mutation: str,
    error_code: str,
) -> None:
    from scripts.release import verify_packaged_gui

    binary_dir, binary_name = _packaged_gui(tmp_path)
    verification_dir = tmp_path / "verification"
    verification_dir.mkdir()
    evidence, evidence_path = _write_office_evidence_fixture(verification_dir)
    cases = evidence["cases"]
    assert isinstance(cases, list)
    first = cases[0]
    assert isinstance(first, dict)
    source = first["source"]
    checks = first["checks"]
    assert isinstance(source, dict)
    assert isinstance(checks, dict)

    if mutation == "wrong-hash":
        source["sha256"] = "0" * 64
    elif mutation == "missing-file":
        source["path"] = "docx/missing.bin"
    elif mutation == "boolean-bytes":
        source["bytes"] = True
    elif mutation == "boolean-page-count":
        checks["pageCount"] = True
    elif mutation == "parent-traversal":
        source["path"] = "../outside.bin"
    else:
        raise AssertionError(mutation)
    verify_packaged_gui._atomic_json_write(evidence_path, evidence)

    with pytest.raises(RuntimeError, match=error_code):
        verify_packaged_gui._build_acceptance_receipt(
            verification_dir,
            candidate_id="candidate-office",
            binary_path=binary_dir / binary_name,
            selected_gates=["office"],
        )


def test_office_gate_summary_rejects_files_changed_after_tree_capture(tmp_path: Path) -> None:
    from scripts.release import verify_packaged_gui

    verification_dir = tmp_path / "verification"
    verification_dir.mkdir()
    evidence, _evidence_path = _write_office_evidence_fixture(verification_dir)
    _directories, files = verify_packaged_gui._capture_evidence_tree(verification_dir)

    cases = evidence["cases"]
    assert isinstance(cases, list)
    first = cases[0]
    assert isinstance(first, dict)
    source = first["source"]
    assert isinstance(source, dict)
    source_path = verification_dir / "gui_office_smoke" / str(source["path"])
    source_path.write_bytes(b"changed after capture")

    with pytest.raises(RuntimeError, match="office_evidence_file_mismatch"):
        verify_packaged_gui._gate_evidence_summary(
            verification_dir,
            selected_gates=["office"],
            files=files,
        )


def test_office_gate_receipt_fails_closed_without_three_case_evidence(tmp_path: Path) -> None:
    from scripts.release import verify_packaged_gui

    binary_dir, binary_name = _packaged_gui(tmp_path)
    verification_dir = tmp_path / "verification"
    office_dir = verification_dir / "gui_office_smoke"
    office_dir.mkdir(parents=True)
    verify_packaged_gui._atomic_json_write(
        office_dir / "office-smoke-evidence.json",
        {
            "schema": verify_packaged_gui._OFFICE_EVIDENCE_SCHEMA,
            "host": {"system": "Windows", "release": "11", "machine": "AMD64", "pythonPlatform": "win32"},
            "cases": [{"case": "docx", "backend": "msoffice_word"}],
        },
    )

    with pytest.raises(RuntimeError, match="office_evidence_shape_invalid"):
        verify_packaged_gui._build_acceptance_receipt(
            verification_dir,
            candidate_id="candidate-office",
            binary_path=binary_dir / binary_name,
            selected_gates=["office"],
        )
