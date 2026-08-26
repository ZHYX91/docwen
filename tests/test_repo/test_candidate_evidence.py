from __future__ import annotations

import hashlib
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

import pytest
from scripts.release import candidate_evidence

pytestmark = pytest.mark.unit


def _git(repo: Path, *args: str) -> None:
    completed = subprocess.run(
        ["git", "-C", str(repo), *args],
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    assert completed.returncode == 0, completed.stderr


def _make_repo(root: Path) -> Path:
    repo = root / "repo"
    repo.mkdir()
    _git(repo, "init", "--quiet")
    _git(repo, "config", "user.name", "DocWen Test")
    _git(repo, "config", "user.email", "docwen-test@example.invalid")
    for name, content in {
        "tracked.txt": "tracked\n",
        "staged.txt": "original\n",
        "deleted.txt": "delete me\n",
    }.items():
        (repo / name).write_text(content, encoding="utf-8")
    _git(repo, "add", ".")
    _git(repo, "commit", "--quiet", "--no-gpg-sign", "-m", "initial")
    return repo


def _make_candidate(root: Path) -> tuple[Path, tuple[str, str]]:
    candidate = root / "candidate"
    packages = ("DocWen_v0.9.0_win-x64", "DocWenCLI_v0.9.0_win-x64")
    for package_name, executable in zip(packages, ("DocWen.exe", "DocWenCLI.exe"), strict=True):
        package = candidate / package_name
        (package / "configs").mkdir(parents=True)
        (package / executable).write_bytes(f"{package_name}-binary".encode())
        (package / "configs" / "配置.toml").write_text("enabled = true\n", encoding="utf-8")
    (candidate / "_evidence").mkdir()
    return candidate, packages


def _directory_link(link: Path, target: Path) -> None:
    if os.name == "nt":
        completed = subprocess.run(
            ["cmd", "/c", "mklink", "/J", str(link), str(target.resolve())],
            check=False,
            capture_output=True,
            text=True,
        )
        assert completed.returncode == 0, completed.stderr or completed.stdout
    else:
        link.symlink_to(target, target_is_directory=True)


def _remove_directory_link(link: Path) -> None:
    """Remove a test link without following its target on either platform."""
    if os.name == "nt":
        link.rmdir()
    else:
        link.unlink()


def test_source_snapshot_captures_head_index_worktree_deletion_and_untracked_files(tmp_path: Path) -> None:
    repo = _make_repo(tmp_path)
    (repo / "tracked.txt").write_text("worktree changed\n", encoding="utf-8")
    (repo / "staged.txt").write_text("staged content\n", encoding="utf-8")
    _git(repo, "add", "staged.txt")
    (repo / "staged.txt").write_text("changed after staging\n", encoding="utf-8")
    (repo / "deleted.txt").unlink()
    (repo / "未跟踪.txt").write_text("证据\n", encoding="utf-8")
    first_output = tmp_path / "first-source.json"
    second_output = tmp_path / "second-source.json"

    first = candidate_evidence.write_source_snapshot(repo, first_output)
    second = candidate_evidence.write_source_snapshot(repo, second_output)

    assert first == second
    assert first["schema"] == candidate_evidence.SOURCE_STATE_SCHEMA
    assert first["git"]["dirty"] is True
    assert len(first["git"]["head"]) == 40
    assert len(first["git"]["headTree"]) == 40
    assert first["sourceStateSha256"] == candidate_evidence._payload_hash(
        {key: value for key, value in first.items() if key != "sourceStateSha256"}
    )
    assert len(first["workingTreeContentSha256"]) == 64
    tracked = {entry["path"]: entry for entry in first["tracked"]["files"]}
    assert tracked["deleted.txt"]["state"] == "missing"
    assert tracked["tracked.txt"]["sha256"] == hashlib.sha256((repo / "tracked.txt").read_bytes()).hexdigest()
    assert tracked["staged.txt"]["sha256"] == hashlib.sha256((repo / "staged.txt").read_bytes()).hexdigest()
    assert tracked["staged.txt"]["indexBlob"] != tracked["tracked.txt"]["indexBlob"]
    untracked_bytes = (repo / "未跟踪.txt").read_bytes()
    assert first["untracked"]["files"] == [
        {
            "path": "未跟踪.txt",
            "size": len(untracked_bytes),
            "sha256": hashlib.sha256(untracked_bytes).hexdigest(),
        }
    ]
    assert json.loads(first_output.read_text(encoding="utf-8")) == first


def test_working_tree_content_identity_is_relocatable_and_changes_with_content(tmp_path: Path) -> None:
    repo = _make_repo(tmp_path)
    (repo / "tracked.txt").write_text("same relocated content\n", encoding="utf-8")
    (repo / "untracked.txt").write_text("same untracked content\n", encoding="utf-8")
    original = candidate_evidence.capture_source_snapshot(repo)
    relocated = tmp_path / "relocated-repo"
    shutil.copytree(repo, relocated)

    relocated_snapshot = candidate_evidence.capture_source_snapshot(relocated)

    assert original["repository"] != relocated_snapshot["repository"]
    assert original["sourceStateSha256"] != relocated_snapshot["sourceStateSha256"]
    assert original["workingTreeContentSha256"] == relocated_snapshot["workingTreeContentSha256"]

    (relocated / "tracked.txt").write_text("different content\n", encoding="utf-8")
    changed = candidate_evidence.capture_source_snapshot(relocated)
    assert changed["workingTreeContentSha256"] != original["workingTreeContentSha256"]


def test_source_snapshot_rejects_existing_or_source_tree_output(tmp_path: Path) -> None:
    repo = _make_repo(tmp_path)
    existing = tmp_path / "existing.json"
    existing.write_text("do not replace", encoding="utf-8")

    with pytest.raises(candidate_evidence.EvidenceError, match="output_exists"):
        candidate_evidence.write_source_snapshot(repo, existing)
    assert existing.read_text(encoding="utf-8") == "do not replace"

    with pytest.raises(candidate_evidence.EvidenceError, match="output_inside_source_repo"):
        candidate_evidence.write_source_snapshot(repo, repo / "unsafe.json")
    assert not (repo / "unsafe.json").exists()


def test_source_snapshot_fails_when_the_tree_changes_between_captures(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    repo = _make_repo(tmp_path)
    original = candidate_evidence._capture_source_state_once
    calls = 0

    def capture_then_change(root: Path) -> dict[str, object]:
        nonlocal calls
        payload = original(root)
        calls += 1
        if calls == 1:
            (root / "tracked.txt").write_text("raced\n", encoding="utf-8")
        return payload

    monkeypatch.setattr(candidate_evidence, "_capture_source_state_once", capture_then_change)

    with pytest.raises(candidate_evidence.EvidenceError, match="source_changed_during_snapshot"):
        candidate_evidence.capture_source_snapshot(repo)


def test_package_manifest_is_deterministic_relocatable_and_detects_drift(tmp_path: Path) -> None:
    candidate, packages = _make_candidate(tmp_path)
    manifest_path = candidate / "_evidence" / "package-manifest.json"

    manifest = candidate_evidence.write_package_manifest(candidate, packages, manifest_path)
    assert manifest == {
        "schemaVersion": 1,
        "algorithm": candidate_evidence.PACKAGE_MANIFEST_SCHEMA,
        "packages": sorted(packages),
        "files": manifest["files"],
    }
    assert len(manifest["files"]) == 4
    assert all(set(item) == {"package", "path", "bytes", "sha256"} for item in manifest["files"])
    assert candidate_evidence.verify_package_manifest(candidate, manifest_path) == manifest

    verification_root = tmp_path / "verify"
    verification_root.mkdir()
    for package in packages:
        shutil.copytree(candidate / package, verification_root / package)
    assert candidate_evidence.verify_package_manifest(verification_root, manifest_path) == manifest

    changed_file = verification_root / packages[0] / "configs" / "配置.toml"
    changed_file.write_text("enabled = false\n", encoding="utf-8")
    with pytest.raises(candidate_evidence.EvidenceError, match="package_manifest_mismatch"):
        candidate_evidence.verify_package_manifest(verification_root, manifest_path)


def test_package_manifest_rejects_unexpected_roots_traversal_links_and_source_tree_candidates(
    tmp_path: Path,
) -> None:
    repo = _make_repo(tmp_path)
    candidate, packages = _make_candidate(tmp_path)
    (candidate / "unexpected.txt").write_text("unexpected", encoding="utf-8")
    with pytest.raises(candidate_evidence.EvidenceError, match="package_root_unexpected_entries"):
        candidate_evidence.capture_package_manifest(candidate, packages)
    (candidate / "unexpected.txt").unlink()

    with pytest.raises(candidate_evidence.EvidenceError, match="invalid_root_entry"):
        candidate_evidence.capture_package_manifest(candidate, ["../escape"])

    external = tmp_path / "external"
    external.mkdir()
    (external / "outside.txt").write_text("outside", encoding="utf-8")
    link = candidate / packages[0] / "linked"
    _directory_link(link, external)
    with pytest.raises(candidate_evidence.EvidenceError, match="linked_path_rejected"):
        candidate_evidence.capture_package_manifest(candidate, packages)
    assert (external / "outside.txt").read_text(encoding="utf-8") == "outside"
    _remove_directory_link(link)

    source_candidate, source_packages = _make_candidate(repo / "generated")
    with pytest.raises(candidate_evidence.EvidenceError, match="candidate_root_inside_source_repo"):
        candidate_evidence.write_package_manifest(
            source_candidate,
            source_packages,
            source_candidate / "_evidence" / "manifest.json",
            source_repo=repo,
        )


def test_package_manifest_fails_when_a_package_changes_between_captures(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    candidate, packages = _make_candidate(tmp_path)
    original = candidate_evidence._capture_package_manifest_once
    calls = 0

    def capture_then_change(root: Path, names: tuple[str, ...], allowed: tuple[str, ...]) -> dict[str, object]:
        nonlocal calls
        payload = original(root, names, allowed)
        calls += 1
        if calls == 1:
            (root / packages[0] / "DocWen.exe").write_bytes(b"changed concurrently")
        return payload

    monkeypatch.setattr(candidate_evidence, "_capture_package_manifest_once", capture_then_change)

    with pytest.raises(candidate_evidence.EvidenceError, match="packages_changed_during_snapshot"):
        candidate_evidence.capture_package_manifest(candidate, packages)


def test_evidence_manifest_excludes_itself_and_detects_changes(tmp_path: Path) -> None:
    evidence = tmp_path / "evidence"
    (evidence / "logs").mkdir(parents=True)
    (evidence / "logs" / "门禁.json").write_text('{"success":true}\n', encoding="utf-8")
    (evidence / "candidate.json").write_text('{"candidate":"C6-V"}\n', encoding="utf-8")
    manifest_path = evidence / "evidence-manifest.json"

    manifest = candidate_evidence.write_evidence_manifest(evidence, manifest_path)

    assert manifest["schema"] == candidate_evidence.EVIDENCE_MANIFEST_SCHEMA
    assert manifest["fileCount"] == 2
    assert {item["path"] for item in manifest["files"]} == {"candidate.json", "logs/门禁.json"}
    assert candidate_evidence.verify_evidence_manifest(evidence, manifest_path) == manifest
    with pytest.raises(candidate_evidence.EvidenceError, match="output_exists"):
        candidate_evidence.write_evidence_manifest(evidence, manifest_path)
    (evidence / "candidate.json").write_text('{"candidate":"changed"}\n', encoding="utf-8")
    with pytest.raises(candidate_evidence.EvidenceError, match="evidence_manifest_mismatch"):
        candidate_evidence.verify_evidence_manifest(evidence, manifest_path)


def test_evidence_manifest_requires_a_new_root_file_outside_the_source_repo(tmp_path: Path) -> None:
    repo = _make_repo(tmp_path)
    evidence = tmp_path / "evidence"
    nested = evidence / "nested"
    nested.mkdir(parents=True)
    (evidence / "log.json").write_text("{}\n", encoding="utf-8")

    with pytest.raises(candidate_evidence.EvidenceError, match="evidence_manifest_must_be_root_file"):
        candidate_evidence.write_evidence_manifest(evidence, nested / "manifest.json")

    source_evidence = repo / "evidence"
    source_evidence.mkdir()
    with pytest.raises(candidate_evidence.EvidenceError, match="evidence_root_inside_source_repo"):
        candidate_evidence.write_evidence_manifest(
            source_evidence,
            source_evidence / "manifest.json",
            source_repo=repo,
        )


def test_evidence_manifest_fails_when_evidence_changes_between_captures(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    evidence = tmp_path / "evidence"
    evidence.mkdir()
    (evidence / "first.json").write_text("{}\n", encoding="utf-8")
    original = candidate_evidence._capture_tree_once
    calls = 0

    def capture_then_change(root: Path, *, excluded_paths: frozenset[str]) -> dict[str, object]:
        nonlocal calls
        payload = original(root, excluded_paths=excluded_paths)
        calls += 1
        if calls == 1:
            (root / "second.json").write_text("{}\n", encoding="utf-8")
        return payload

    monkeypatch.setattr(candidate_evidence, "_capture_tree_once", capture_then_change)

    with pytest.raises(candidate_evidence.EvidenceError, match="tree_changed_during_snapshot"):
        candidate_evidence.capture_evidence_manifest(evidence, excluded_relative_path="manifest.json")


def test_candidate_evidence_self_test_and_cli_are_directly_executable(tmp_path: Path) -> None:
    payload = candidate_evidence.run_self_test()
    assert payload["schema"] == candidate_evidence.SELF_TEST_SCHEMA
    assert payload["success"] is True
    assert all(len(payload[key]) == 64 for key in payload if key.endswith("Sha256"))
    assert "packageManifestFileSha256" in payload
    assert "evidenceManifestFileSha256" in payload
    assert "evidenceContentManifestSha256" in payload

    completed = subprocess.run(
        [sys.executable, "scripts/release/candidate_evidence.py", "self-test"],
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        cwd=Path.cwd(),
    )
    assert completed.returncode == 0, completed.stderr
    assert json.loads(completed.stdout)["data"]["success"] is True


def test_cli_reports_raw_package_and_evidence_manifest_file_hashes(tmp_path: Path) -> None:
    source_repo = _make_repo(tmp_path)
    candidate = tmp_path / "candidate"
    package = candidate / "DocWen_test"
    evidence = candidate / "_evidence"
    package.mkdir(parents=True)
    evidence.mkdir()
    (package / "DocWen.exe").write_bytes(b"candidate")
    package_manifest = evidence / "package-manifest.json"
    package_command = subprocess.run(
        [
            sys.executable,
            "scripts/release/candidate_evidence.py",
            "package-manifest",
            "--repo",
            str(source_repo),
            "--candidate-root",
            str(candidate),
            "--package",
            package.name,
            "--output",
            str(package_manifest),
        ],
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    assert package_command.returncode == 0, package_command.stderr
    package_data = json.loads(package_command.stdout)["data"]
    assert package_data["packageManifestFileSha256"] == hashlib.sha256(package_manifest.read_bytes()).hexdigest()

    evidence_manifest = evidence / "evidence-manifest.json"
    evidence_command = subprocess.run(
        [
            sys.executable,
            "scripts/release/candidate_evidence.py",
            "evidence-manifest",
            "--repo",
            str(source_repo),
            "--evidence-root",
            str(evidence),
            "--output",
            str(evidence_manifest),
        ],
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    assert evidence_command.returncode == 0, evidence_command.stderr
    evidence_data = json.loads(evidence_command.stdout)["data"]
    assert evidence_data["evidenceManifestFileSha256"] == hashlib.sha256(evidence_manifest.read_bytes()).hexdigest()
    assert evidence_data["manifestSha256"] != evidence_data["evidenceManifestFileSha256"]
