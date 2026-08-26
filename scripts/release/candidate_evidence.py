from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import stat
import subprocess
import sys
import tempfile
from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any, NoReturn, cast

SOURCE_STATE_SCHEMA = "docwen-source-state-v1"
SOURCE_CONTENT_SCHEMA = "docwen-source-content-manifest-v1"
UNTRACKED_SOURCE_SCHEMA = "docwen-untracked-source-manifest-v1"
PACKAGE_MANIFEST_SCHEMA = "docwen-package-manifest-v1"
EVIDENCE_MANIFEST_SCHEMA = "docwen-evidence-manifest-v1"
SELF_TEST_SCHEMA = "docwen-candidate-evidence-self-test-v1"

_REPARSE_ATTRIBUTE = getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0x0400)
_READ_CHUNK_SIZE = 1024 * 1024


class EvidenceError(RuntimeError):
    """A fail-closed candidate evidence validation error."""


def _fail(code: str, detail: str | None = None) -> NoReturn:
    if detail:
        raise EvidenceError(f"{code}: {detail}")
    raise EvidenceError(code)


def _canonical_json_bytes(value: object) -> bytes:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")


def _sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def _payload_hash(value: object) -> str:
    return _sha256_bytes(_canonical_json_bytes(value))


def _is_reparse(stats: os.stat_result) -> bool:
    return bool(getattr(stats, "st_file_attributes", 0) & _REPARSE_ATTRIBUTE)


def _path_signature(stats: os.stat_result) -> tuple[int, int, int, int, int]:
    return (
        stats.st_mode,
        stats.st_size,
        stats.st_mtime_ns,
        stats.st_dev,
        stats.st_ino,
    )


def _opened_file_signature(stats: os.stat_result) -> tuple[int, int, int, int]:
    return (
        stats.st_size,
        stats.st_mtime_ns,
        stats.st_dev,
        stats.st_ino,
    )


def _absolute_path(path: Path) -> Path:
    return Path(os.path.abspath(os.fspath(path)))


def _assert_existing_chain_without_links(path: Path) -> None:
    absolute = _absolute_path(path)
    current = Path(absolute.anchor)
    for part in absolute.parts[1:]:
        current /= part
        try:
            stats = current.lstat()
        except FileNotFoundError:
            _fail("path_missing", str(current))
        if stat.S_ISLNK(stats.st_mode) or _is_reparse(stats):
            _fail("linked_path_rejected", str(current))


def _safe_existing_directory(path: Path, *, label: str) -> Path:
    absolute = _absolute_path(path)
    _assert_existing_chain_without_links(absolute)
    stats = absolute.lstat()
    if not stat.S_ISDIR(stats.st_mode):
        _fail("directory_required", f"{label}={absolute}")
    return absolute.resolve(strict=True)


def _safe_regular_file(path: Path, *, label: str) -> Path:
    absolute = _absolute_path(path)
    _assert_existing_chain_without_links(absolute)
    stats = absolute.lstat()
    if not stat.S_ISREG(stats.st_mode):
        _fail("regular_file_required", f"{label}={absolute}")
    return absolute.resolve(strict=True)


def _safe_new_output(path: Path, *, source_repo: Path | None = None) -> Path:
    absolute = _absolute_path(path)
    parent = _safe_existing_directory(absolute.parent, label="output_parent")
    output = parent / absolute.name
    if output.exists() or output.is_symlink():
        _fail("output_exists", str(output))
    if source_repo is not None and _is_same_or_descendant(output, source_repo):
        _fail("output_inside_source_repo", str(output))
    return output


def _is_same_or_descendant(candidate: Path, root: Path) -> bool:
    try:
        candidate.relative_to(root)
    except ValueError:
        return False
    return True


def _write_json_exclusive(path: Path, payload: object) -> None:
    with path.open("x", encoding="utf-8", newline="\n") as handle:
        json.dump(payload, handle, ensure_ascii=False, sort_keys=True, indent=2)
        handle.write("\n")
        handle.flush()
        os.fsync(handle.fileno())


def _read_json_object(path: Path, *, label: str) -> dict[str, Any]:
    safe_path = _safe_regular_file(path, label=label)
    try:
        value = json.loads(safe_path.read_text(encoding="utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        _fail("invalid_json", f"{safe_path}: {exc}")
    if not isinstance(value, dict):
        _fail("json_object_required", str(safe_path))
    return value


def _git(repo: Path, *args: str, allowed_returncodes: tuple[int, ...] = (0,)) -> bytes:
    completed = subprocess.run(
        ["git", "-C", str(repo), *args],
        check=False,
        capture_output=True,
    )
    if completed.returncode not in allowed_returncodes:
        stderr = completed.stderr.decode("utf-8", errors="replace").strip()
        _fail("git_command_failed", f"git {' '.join(args)}: {stderr or completed.returncode}")
    return completed.stdout


def _resolve_git_repo(repo: Path) -> Path:
    requested = _safe_existing_directory(repo, label="source_repo")
    root_text = _git(requested, "rev-parse", "--show-toplevel").decode("utf-8", errors="strict").strip()
    root = _safe_existing_directory(Path(root_text), label="git_root")
    if requested != root:
        _fail("repo_must_be_git_top_level", f"requested={requested}; root={root}")
    return root


def _decode_git_path(value: bytes) -> str:
    try:
        return value.decode("utf-8", errors="strict").replace("\\", "/")
    except UnicodeDecodeError as exc:
        _fail("git_path_not_utf8", str(exc))


def _split_nul(value: bytes) -> list[bytes]:
    if not value:
        return []
    parts = value.split(b"\0")
    if parts[-1] != b"":
        _fail("git_output_missing_nul_terminator")
    return parts[:-1]


def _hash_regular_file(path: Path) -> tuple[int, str]:
    before = path.lstat()
    if stat.S_ISLNK(before.st_mode) or _is_reparse(before):
        _fail("linked_path_rejected", str(path))
    if not stat.S_ISREG(before.st_mode):
        _fail("regular_file_required", str(path))

    digest = hashlib.sha256()
    with path.open("rb") as handle:
        opened = os.fstat(handle.fileno())
        if not stat.S_ISREG(opened.st_mode) or (opened.st_dev, opened.st_ino) != (before.st_dev, before.st_ino):
            _fail("file_changed_during_read", str(path))
        while chunk := handle.read(_READ_CHUNK_SIZE):
            digest.update(chunk)
        after_open = os.fstat(handle.fileno())
    try:
        after_path = path.lstat()
    except FileNotFoundError:
        _fail("file_changed_during_read", str(path))
    if _opened_file_signature(before) != _opened_file_signature(after_open) or _path_signature(
        before
    ) != _path_signature(after_path):
        _fail("file_changed_during_read", str(path))
    return before.st_size, digest.hexdigest()


def _worktree_file(repo: Path, relative_path: str) -> dict[str, object]:
    path = repo.joinpath(*relative_path.split("/"))
    try:
        stats = path.lstat()
    except FileNotFoundError:
        return {"state": "missing", "size": None, "sha256": None}
    if stat.S_ISLNK(stats.st_mode) or _is_reparse(stats):
        _fail("linked_path_rejected", str(path))
    if not stat.S_ISREG(stats.st_mode):
        _fail("tracked_path_not_regular", str(path))
    size, digest = _hash_regular_file(path)
    return {"state": "file", "size": size, "sha256": digest}


def _tracked_manifest(repo: Path) -> dict[str, object]:
    entries: list[dict[str, object]] = []
    for raw_entry in _split_nul(_git(repo, "ls-files", "--stage", "-z")):
        try:
            metadata, raw_path = raw_entry.split(b"\t", 1)
            raw_mode, raw_blob, raw_stage = metadata.split(b" ", 2)
        except ValueError:
            _fail("unexpected_git_ls_files_entry", repr(raw_entry))
        relative_path = _decode_git_path(raw_path)
        entry: dict[str, object] = {
            "path": relative_path,
            "indexMode": raw_mode.decode("ascii"),
            "indexBlob": raw_blob.decode("ascii"),
            "stage": int(raw_stage.decode("ascii")),
        }
        entry.update(_worktree_file(repo, relative_path))
        entries.append(entry)
    entries.sort(key=lambda item: (str(item["path"]), cast(int, item["stage"])))
    payload: dict[str, object] = {
        "schema": SOURCE_CONTENT_SCHEMA,
        "files": entries,
        "fileCount": len(entries),
        "totalBytes": sum(cast(int, entry["size"]) for entry in entries if entry["size"] is not None),
    }
    payload["manifestSha256"] = _payload_hash(entries)
    return payload


def _untracked_manifest(repo: Path) -> dict[str, object]:
    entries: list[dict[str, object]] = []
    raw_paths = _split_nul(_git(repo, "ls-files", "--others", "--exclude-standard", "-z"))
    for raw_path in raw_paths:
        relative_path = _decode_git_path(raw_path)
        path = repo.joinpath(*relative_path.split("/"))
        size, digest = _hash_regular_file(path)
        entries.append({"path": relative_path, "size": size, "sha256": digest})
    entries.sort(key=lambda item: str(item["path"]))
    payload: dict[str, object] = {
        "schema": UNTRACKED_SOURCE_SCHEMA,
        "files": entries,
        "fileCount": len(entries),
        "totalBytes": sum(cast(int, entry["size"]) for entry in entries),
    }
    payload["manifestSha256"] = _payload_hash(entries)
    return payload


def _capture_source_state_once(repo: Path) -> dict[str, object]:
    head = _git(repo, "rev-parse", "HEAD").decode("ascii").strip()
    tree = _git(repo, "rev-parse", "HEAD^{tree}").decode("ascii").strip()
    branch_result = subprocess.run(
        ["git", "-C", str(repo), "symbolic-ref", "--quiet", "--short", "HEAD"],
        check=False,
        capture_output=True,
    )
    if branch_result.returncode not in (0, 1):
        _fail("git_command_failed", branch_result.stderr.decode("utf-8", errors="replace").strip())
    branch = branch_result.stdout.decode("utf-8", errors="strict").strip() if branch_result.returncode == 0 else None
    status = _git(repo, "status", "--porcelain=v2", "-z", "--untracked-files=all")
    return {
        "schema": SOURCE_STATE_SCHEMA,
        "repository": str(repo),
        "git": {
            "branch": branch,
            "head": head,
            "headTree": tree,
            "dirty": bool(status),
            "statusPorcelainV2Bytes": len(status),
            "statusPorcelainV2Sha256": _sha256_bytes(status),
        },
        "tracked": _tracked_manifest(repo),
        "untracked": _untracked_manifest(repo),
    }


def capture_source_snapshot(repo: Path) -> dict[str, object]:
    root = _resolve_git_repo(repo)
    first = _capture_source_state_once(root)
    second = _capture_source_state_once(root)
    if first != second:
        _fail("source_changed_during_snapshot")
    first["workingTreeContentSha256"] = _payload_hash({"tracked": first["tracked"], "untracked": first["untracked"]})
    first["sourceStateSha256"] = _payload_hash(first)
    return first


def write_source_snapshot(repo: Path, output: Path) -> dict[str, object]:
    root = _resolve_git_repo(repo)
    safe_output = _safe_new_output(output, source_repo=root)
    snapshot = capture_source_snapshot(root)
    _write_json_exclusive(safe_output, snapshot)
    return snapshot


def _validate_root_entry_name(name: str, *, label: str) -> str:
    if not name or name in (".", "..") or Path(name).name != name or "/" in name or "\\" in name:
        _fail("invalid_root_entry", f"{label}={name!r}")
    return name


def _safe_root_entries(root: Path) -> set[str]:
    names: set[str] = set()
    with os.scandir(root) as iterator:
        for entry in iterator:
            stats = entry.stat(follow_symlinks=False)
            if entry.is_symlink() or _is_reparse(stats):
                _fail("linked_path_rejected", str(root / entry.name))
            names.add(entry.name)
    return names


def _capture_tree_once(root: Path, *, excluded_paths: frozenset[str] = frozenset()) -> dict[str, object]:
    entries: list[dict[str, object]] = []

    def visit(directory: Path, relative_directory: str) -> None:
        with os.scandir(directory) as iterator:
            children = sorted(iterator, key=lambda child: child.name)
        for child in children:
            relative_path = f"{relative_directory}/{child.name}" if relative_directory else child.name
            normalized = relative_path.replace("\\", "/")
            if normalized in excluded_paths:
                continue
            stats = child.stat(follow_symlinks=False)
            child_path = directory / child.name
            if child.is_symlink() or _is_reparse(stats):
                _fail("linked_path_rejected", str(child_path))
            if stat.S_ISDIR(stats.st_mode):
                visit(child_path, normalized)
                continue
            if not stat.S_ISREG(stats.st_mode):
                _fail("unsupported_file_type", str(child_path))
            size, digest = _hash_regular_file(child_path)
            entries.append({"path": normalized, "size": size, "sha256": digest})

    visit(root, "")
    entries.sort(key=lambda item: str(item["path"]))
    payload: dict[str, object] = {
        "files": entries,
        "fileCount": len(entries),
        "totalBytes": sum(cast(int, entry["size"]) for entry in entries),
    }
    payload["manifestSha256"] = _payload_hash(entries)
    return payload


def _capture_tree_stable(root: Path, *, excluded_paths: frozenset[str] = frozenset()) -> dict[str, object]:
    first = _capture_tree_once(root, excluded_paths=excluded_paths)
    second = _capture_tree_once(root, excluded_paths=excluded_paths)
    if first != second:
        _fail("tree_changed_during_snapshot", str(root))
    return first


def _capture_package_manifest_once(
    candidate_root: Path,
    package_names: Sequence[str],
    allowed_root_entries: Sequence[str],
) -> dict[str, object]:
    packages = tuple(_validate_root_entry_name(name, label="package") for name in package_names)
    if not packages or len(set(packages)) != len(packages):
        _fail("unique_packages_required")
    allowed = tuple(_validate_root_entry_name(name, label="allowed_root_entry") for name in allowed_root_entries)
    if len(set(allowed)) != len(allowed) or set(packages) & set(allowed):
        _fail("invalid_allowed_root_entries")

    observed = _safe_root_entries(candidate_root)
    missing = sorted(set(packages) - observed)
    unexpected = sorted(observed - set(packages) - set(allowed))
    if missing:
        _fail("package_root_missing_entries", repr(missing))
    if unexpected:
        _fail("package_root_unexpected_entries", repr(unexpected))
    for allowed_name in set(allowed) & observed:
        allowed_path = candidate_root / allowed_name
        if not stat.S_ISDIR(allowed_path.lstat().st_mode):
            _fail("allowed_root_entry_not_directory", str(allowed_path))

    files: list[dict[str, object]] = []
    for package_name in sorted(packages):
        package_root = _safe_existing_directory(candidate_root / package_name, label=f"package:{package_name}")
        tree = _capture_tree_once(package_root)
        for entry in cast(list[dict[str, object]], tree["files"]):
            files.append(
                {
                    "package": package_name,
                    "path": entry["path"],
                    "bytes": entry["size"],
                    "sha256": entry["sha256"],
                }
            )
    files.sort(key=lambda item: (str(item["package"]), str(item["path"])))
    if not files:
        _fail("package_manifest_has_no_files")
    return {
        "schemaVersion": 1,
        "algorithm": PACKAGE_MANIFEST_SCHEMA,
        "packages": sorted(packages),
        "files": files,
    }


def capture_package_manifest(
    candidate_root: Path,
    package_names: Sequence[str],
    *,
    allowed_root_entries: Sequence[str] = ("_evidence",),
) -> dict[str, object]:
    root = _safe_existing_directory(candidate_root, label="candidate_root")
    first = _capture_package_manifest_once(root, package_names, allowed_root_entries)
    second = _capture_package_manifest_once(root, package_names, allowed_root_entries)
    if first != second:
        _fail("packages_changed_during_snapshot")
    return first


def write_package_manifest(
    candidate_root: Path,
    package_names: Sequence[str],
    output: Path,
    *,
    allowed_root_entries: Sequence[str] = ("_evidence",),
    source_repo: Path | None = None,
) -> dict[str, object]:
    root = _safe_existing_directory(candidate_root, label="candidate_root")
    repo = _resolve_git_repo(source_repo) if source_repo is not None else None
    if repo is not None and _is_same_or_descendant(root, repo):
        _fail("candidate_root_inside_source_repo", str(root))
    safe_output = _safe_new_output(output, source_repo=repo)
    manifest = capture_package_manifest(root, package_names, allowed_root_entries=allowed_root_entries)
    _write_json_exclusive(safe_output, manifest)
    return manifest


def _package_manifest_contract(manifest: Mapping[str, Any]) -> list[str]:
    if manifest.get("schemaVersion") != 1 or manifest.get("algorithm") != PACKAGE_MANIFEST_SCHEMA:
        _fail("package_manifest_schema_mismatch")
    raw_packages = manifest.get("packages")
    raw_files = manifest.get("files")
    if not isinstance(raw_packages, list) or not isinstance(raw_files, list) or not raw_files:
        _fail("invalid_package_manifest")
    if not raw_packages or not all(isinstance(name, str) for name in raw_packages):
        _fail("invalid_package_manifest")
    package_names = [_validate_root_entry_name(cast(str, name), label="package") for name in raw_packages]
    if len(set(package_names)) != len(package_names):
        _fail("invalid_package_manifest")
    keys: set[tuple[str, str]] = set()
    for raw_entry in raw_files:
        if not isinstance(raw_entry, dict) or set(raw_entry) != {"package", "path", "bytes", "sha256"}:
            _fail("invalid_package_manifest")
        entry = cast(dict[str, Any], raw_entry)
        package = entry["package"]
        relative_path = entry["path"]
        byte_count = entry["bytes"]
        digest = entry["sha256"]
        if (
            not isinstance(package, str)
            or package not in package_names
            or not isinstance(relative_path, str)
            or not relative_path
            or Path(relative_path).is_absolute()
            or relative_path == ".."
            or relative_path.startswith("../")
            or "\\" in relative_path
            or any(part in ("", ".", "..") for part in relative_path.split("/"))
            or not isinstance(byte_count, int)
            or isinstance(byte_count, bool)
            or byte_count < 0
            or not isinstance(digest, str)
            or len(digest) != 64
            or any(character not in "0123456789abcdef" for character in digest)
        ):
            _fail("invalid_package_manifest")
        key = (package, relative_path)
        if key in keys:
            _fail("invalid_package_manifest")
        keys.add(key)
    return package_names


def verify_package_manifest(
    candidate_root: Path,
    manifest_path: Path,
    *,
    allowed_root_entries: Sequence[str] = ("_evidence",),
) -> dict[str, object]:
    manifest = _read_json_object(manifest_path, label="package_manifest")
    package_names = _package_manifest_contract(manifest)
    actual = capture_package_manifest(candidate_root, package_names, allowed_root_entries=allowed_root_entries)
    if actual != manifest:
        _fail("package_manifest_mismatch")
    return actual


def capture_evidence_manifest(evidence_root: Path, *, excluded_relative_path: str) -> dict[str, object]:
    root = _safe_existing_directory(evidence_root, label="evidence_root")
    excluded = frozenset({excluded_relative_path.replace("\\", "/")})
    tree = _capture_tree_stable(root, excluded_paths=excluded)
    return {"schema": EVIDENCE_MANIFEST_SCHEMA, **tree}


def write_evidence_manifest(
    evidence_root: Path,
    output: Path,
    *,
    source_repo: Path | None = None,
) -> dict[str, object]:
    root = _safe_existing_directory(evidence_root, label="evidence_root")
    repo = _resolve_git_repo(source_repo) if source_repo is not None else None
    if repo is not None and _is_same_or_descendant(root, repo):
        _fail("evidence_root_inside_source_repo", str(root))
    safe_output = _safe_new_output(output, source_repo=repo)
    if safe_output.parent != root:
        _fail("evidence_manifest_must_be_root_file", str(safe_output))
    manifest = capture_evidence_manifest(root, excluded_relative_path=safe_output.name)
    _write_json_exclusive(safe_output, manifest)
    return manifest


def verify_evidence_manifest(evidence_root: Path, manifest_path: Path) -> dict[str, object]:
    root = _safe_existing_directory(evidence_root, label="evidence_root")
    manifest_file = _safe_regular_file(manifest_path, label="evidence_manifest")
    if manifest_file.parent != root:
        _fail("evidence_manifest_must_be_root_file", str(manifest_file))
    manifest = _read_json_object(manifest_file, label="evidence_manifest")
    if manifest.get("schema") != EVIDENCE_MANIFEST_SCHEMA:
        _fail("evidence_manifest_schema_mismatch")
    actual = capture_evidence_manifest(root, excluded_relative_path=manifest_file.name)
    if actual != manifest:
        _fail(
            "evidence_manifest_mismatch",
            f"expected={manifest.get('manifestSha256')}; actual={actual.get('manifestSha256')}",
        )
    return actual


def _run_git(repo: Path, *args: str) -> None:
    completed = subprocess.run(["git", "-C", str(repo), *args], check=False, capture_output=True)
    if completed.returncode != 0:
        _fail("self_test_git_failed", completed.stderr.decode("utf-8", errors="replace").strip())


def run_self_test() -> dict[str, object]:
    with tempfile.TemporaryDirectory(prefix="docwen-candidate-evidence-") as temporary:
        root = Path(temporary).resolve()
        repo = root / "repo"
        repo.mkdir()
        _run_git(repo, "init", "--quiet")
        _run_git(repo, "config", "user.name", "DocWen Evidence Self Test")
        _run_git(repo, "config", "user.email", "evidence-self-test@example.invalid")
        (repo / "tracked.txt").write_text("tracked\n", encoding="utf-8")
        _run_git(repo, "add", "tracked.txt")
        _run_git(repo, "commit", "--quiet", "--no-gpg-sign", "-m", "self test")
        (repo / "untracked-资料.txt").write_text("证据\n", encoding="utf-8")
        source_output = root / "source.json"
        source = write_source_snapshot(repo, source_output)

        candidate = root / "candidate"
        package = candidate / "DocWen_self_test"
        evidence = candidate / "_evidence"
        package.mkdir(parents=True)
        evidence.mkdir()
        (package / "DocWen.exe").write_bytes(b"self-test-package")
        package_manifest_path = evidence / "package-manifest.json"
        write_package_manifest(
            candidate,
            [package.name],
            package_manifest_path,
            source_repo=repo,
        )
        verify_package_manifest(candidate, package_manifest_path)
        shutil.copy2(source_output, evidence / "source.json")
        evidence_manifest_path = evidence / "evidence-manifest.json"
        evidence_manifest = write_evidence_manifest(evidence, evidence_manifest_path, source_repo=repo)
        verify_evidence_manifest(evidence, evidence_manifest_path)
        return {
            "schema": SELF_TEST_SCHEMA,
            "success": True,
            "sourceStateSha256": source["sourceStateSha256"],
            "packageManifestFileSha256": _sha256_bytes(package_manifest_path.read_bytes()),
            "evidenceManifestFileSha256": _sha256_bytes(evidence_manifest_path.read_bytes()),
            "evidenceContentManifestSha256": evidence_manifest["manifestSha256"],
        }


def _add_repo_argument(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--repo", type=Path, default=Path.cwd(), help="Exact Git repository root (default: cwd).")


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Create and verify immutable local DocWen candidate evidence.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    source = subparsers.add_parser("source-snapshot", help="Capture a stable Git/index/worktree source snapshot.")
    _add_repo_argument(source)
    source.add_argument("--output", type=Path, required=True)

    package = subparsers.add_parser("package-manifest", help="Create a deterministic manifest for exact package roots.")
    _add_repo_argument(package)
    package.add_argument("--candidate-root", type=Path, required=True)
    package.add_argument("--package", action="append", required=True, dest="packages")
    package.add_argument("--allow-root-entry", action="append", default=[], dest="allowed_root_entries")
    package.add_argument("--output", type=Path, required=True)

    package_verify = subparsers.add_parser("verify-package", help="Verify package bytes against a saved manifest.")
    package_verify.add_argument("--candidate-root", type=Path, required=True)
    package_verify.add_argument("--manifest", type=Path, required=True)

    evidence = subparsers.add_parser("evidence-manifest", help="Create a self-excluding evidence manifest.")
    _add_repo_argument(evidence)
    evidence.add_argument("--evidence-root", type=Path, required=True)
    evidence.add_argument("--output", type=Path, required=True)

    evidence_verify = subparsers.add_parser("verify-evidence", help="Verify an evidence directory manifest.")
    evidence_verify.add_argument("--evidence-root", type=Path, required=True)
    evidence_verify.add_argument("--manifest", type=Path, required=True)

    subparsers.add_parser("self-test", help="Run an isolated end-to-end evidence-tool self-test.")
    return parser


def _success(command: str, payload: Mapping[str, object]) -> None:
    print(json.dumps({"success": True, "command": command, "data": payload}, ensure_ascii=False, sort_keys=True))


def main(argv: list[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)
    try:
        if args.command == "source-snapshot":
            payload = write_source_snapshot(args.repo, args.output)
            _success(args.command, {"output": str(_absolute_path(args.output)), **payload})
        elif args.command == "package-manifest":
            allowed = args.allowed_root_entries or ["_evidence"]
            payload = write_package_manifest(
                args.candidate_root,
                args.packages,
                args.output,
                allowed_root_entries=allowed,
                source_repo=args.repo,
            )
            _success(
                args.command,
                {
                    "output": str(_absolute_path(args.output)),
                    "packageManifestFileSha256": _sha256_bytes(_absolute_path(args.output).read_bytes()),
                    **payload,
                },
            )
        elif args.command == "verify-package":
            payload = verify_package_manifest(args.candidate_root, args.manifest)
            _success(args.command, payload)
        elif args.command == "evidence-manifest":
            payload = write_evidence_manifest(args.evidence_root, args.output, source_repo=args.repo)
            _success(
                args.command,
                {
                    "output": str(_absolute_path(args.output)),
                    "evidenceManifestFileSha256": _sha256_bytes(_absolute_path(args.output).read_bytes()),
                    **payload,
                },
            )
        elif args.command == "verify-evidence":
            payload = verify_evidence_manifest(args.evidence_root, args.manifest)
            _success(args.command, payload)
        elif args.command == "self-test":
            _success(args.command, run_self_test())
        else:
            _fail("unknown_command", str(args.command))
    except EvidenceError as exc:
        print(f"candidate_evidence_error: {exc}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
