#!/usr/bin/env python3
"""Build clean DocWen v4 package/evidence inputs without sealing a candidate.

This producer is intentionally upstream of ``build_v4_candidate_staging.py``.
It builds from detached disposable clones, runs the frozen exact-two Machine
harness, and emits only the package input, evidence input, and strict evidence
plan consumed by the existing staging/sealing tools.  Host evidence is never
invented here: Word, WPS, and LibreOffice remain ``not_run``.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
import uuid
from collections.abc import Callable, Sequence
from pathlib import Path
from typing import Any

if __package__ in {None, ""}:
    _BOOTSTRAP_ROOT = Path(__file__).resolve().parents[2]
    if str(_BOOTSTRAP_ROOT) not in sys.path:
        sys.path.insert(0, str(_BOOTSTRAP_ROOT))

from scripts.release import build_v4_candidate_staging as staging_builder
from scripts.release import candidate_evidence
from scripts.release import v4_candidate_contract as contract
from scripts.release import v4_evidence_contract as evidence_contract
from scripts.release import v4_evidence_io as evidence_io
from scripts.release import v4_package_input_build as package_build
from scripts.release import v4_package_input_contract as input_contract
from scripts.release import v4_package_input_evidence as input_evidence
from scripts.release import v4_package_input_runner as machine_runner

HARNESS_ID = input_contract.HARNESS_ID
HARNESS_VERSION = input_contract.HARNESS_VERSION
HARNESS_SCHEMA = input_contract.HARNESS_SCHEMA
HARNESS_MANIFEST_RELATIVE = input_contract.HARNESS_MANIFEST_RELATIVE
REQUIRED_HARNESS_CASE_IDS = input_contract.REQUIRED_HARNESS_CASE_IDS
PRODUCER_RESULT_SCHEMA = "docwen.v4_package_input_result.v1"
PLAN_SCHEMA = staging_builder.PLAN_SCHEMA
PACKAGE_NAMES = staging_builder.PACKAGE_NAMES
SOURCE_ORACLE_ROOT = input_contract.SOURCE_ORACLE_ROOT
SOURCE_ORACLE_MANIFEST = SOURCE_ORACLE_ROOT / "manifest.json"
VALIDATION_CASE_ID = input_contract.VALIDATION_CASE_ID
HARNESS_CASE_PREFIX = input_contract.HARNESS_CASE_PREFIX
EXACT_CAPABILITY_ID = input_contract.EXACT_CAPABILITY_ID
NEUTRAL_MEDIA_TYPE = input_contract.NEUTRAL_MEDIA_TYPE
PLAN_MEDIA_TYPE = input_contract.PLAN_MEDIA_TYPE
DOCX_MEDIA_TYPE = input_contract.DOCX_MEDIA_TYPE
V4PackageInputError = input_contract.V4PackageInputError
BuildOutput = input_contract.BuildOutput
HarnessInput = input_contract.HarnessInput
HarnessOutput = input_contract.HarnessOutput
HarnessCaseOutput = input_contract.HarnessCaseOutput

CheckpointLoader = Callable[[Path, str, Path], tuple[dict[str, Any], dict[str, object]]]
CloneFactory = Callable[[Path, Path, str, str, str], None]
CloneVerifier = Callable[[Path, str, str, str], None]
PackageBuilder = Callable[[Path, Path, Path], BuildOutput]
HarnessRunner = Callable[[Path, Path, Path, HarnessInput], HarnessOutput]
VersionReader = Callable[[Path], str]


def _sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def _write_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    contract.write_json_exclusive(path, value)


def _identity(path: Path, *, root: Path) -> dict[str, object]:
    try:
        return evidence_io.file_identity(path, relative_to=root)
    except evidence_io.V4EvidenceContractError as exc:
        raise V4PackageInputError(str(exc)) from exc


def _require_new_external_root(path: Path, *, label: str, repositories: Sequence[Path]) -> Path:
    absolute = Path(os.path.abspath(path))
    if not absolute.name or ":" in absolute.name or "\\" in absolute.name:
        raise V4PackageInputError(f"{label}_path_invalid:{absolute}")
    try:
        parent = candidate_evidence._safe_existing_directory(absolute.parent, label=f"{label}_parent")
    except candidate_evidence.EvidenceError as exc:
        raise V4PackageInputError(str(exc)) from exc
    target = parent / absolute.name
    if target.exists() or target.is_symlink():
        raise V4PackageInputError(f"{label}_already_exists:{target}")
    for raw in repositories:
        repository = raw.resolve(strict=True)
        if target == repository or repository in target.parents or target in repository.parents:
            raise V4PackageInputError(f"{label}_overlaps_repository:{repository}")
    return target


_require_command = input_contract.require_command
_load_checkpoint = input_contract.load_checkpoint
_clone_exact = input_contract.clone_exact
_verify_clone_identity = input_contract.verify_clone_identity
_default_package_builder = package_build.default_package_builder


_reject_legacy_harness = input_contract.reject_legacy_harness
_load_harness = input_contract.load_harness
_revalidate_harness = input_contract.revalidate_harness
_exact_two_inputs = input_contract.exact_two_inputs
_validate_exact_two_request = input_contract.validate_exact_two_request
_transcript_event = input_contract.transcript_event
_transcript_request_digest = input_contract.transcript_request_digest
_build_session_transcript = input_contract.build_session_transcript
_validate_session_transcript = input_contract.validate_session_transcript
_inspect_docx = input_contract.inspect_docx


_frame = machine_runner._frame
_default_harness_runner = machine_runner._default_harness_runner


def _default_version_reader(executable: Path) -> str:
    completed = _require_command([str(executable), "--version"], cwd=executable.parent, label="binary_version")
    if completed.stderr:
        raise V4PackageInputError("binary_version_stderr_not_empty")
    try:
        value = completed.stdout.decode("utf-8", errors="strict").strip()
    except UnicodeDecodeError as exc:
        raise V4PackageInputError("binary_version_not_utf8") from exc
    expected = f"DocWen {contract.PRODUCT_VERSION} (CLI protocol 3)"
    if value != expected:
        raise V4PackageInputError(f"binary_version_mismatch:{value}")
    return value


def _copy_tree_stable(source: Path, destination: Path) -> None:
    try:
        before = evidence_contract.capture_tree_stable(source)
        shutil.copytree(source, destination)
        after_source = evidence_contract.capture_tree_stable(source)
        after_destination = evidence_contract.capture_tree_stable(destination)
    except evidence_contract.V4EvidenceContractError as exc:
        raise V4PackageInputError(str(exc)) from exc
    if before != after_source or before != after_destination:
        raise V4PackageInputError(f"package_copy_not_stable:{source.name}")


_build_evidence = input_evidence.build_evidence


def build_package_input(
    *,
    docwen_repo: Path,
    source_checkpoint: Path,
    source_checkpoint_sha256: str,
    candidate_id: str,
    python: Path,
    uv: Path,
    output_root: Path,
    work_root: Path,
    checkpoint_loader: CheckpointLoader = _load_checkpoint,
    clone_factory: CloneFactory = _clone_exact,
    clone_verifier: CloneVerifier = _verify_clone_identity,
    package_builder: PackageBuilder = _default_package_builder,
    harness_runner: HarnessRunner = _default_harness_runner,
    version_reader: VersionReader = _default_version_reader,
) -> dict[str, object]:
    # The reverse-neutral capability is gated by the frozen v2 harness: a
    # harness that still lists pending case IDs is rejected by
    # ``load_harness`` before any package or evidence is produced.  The
    # default runner additionally remains fail-closed until the packaged
    # DocWen CLI is available and the v2 exact-neutral recovery authority is
    # bound to every executed case.  This source-level producer never
    # substitutes legacy `convert.docx.to_markdown` output for the
    # authenticated recovery proof.
    repositories = (docwen_repo.resolve(strict=True),)
    output = _require_new_external_root(output_root, label="output_root", repositories=repositories)
    work = _require_new_external_root(work_root, label="work_root", repositories=repositories)
    if output == work or output in work.parents or work in output.parents:
        raise V4PackageInputError("output_and_work_roots_overlap")
    checkpoint, checkpoint_identity = checkpoint_loader(
        source_checkpoint,
        source_checkpoint_sha256,
        repositories[0],
    )
    docwen = checkpoint.get("docwen")
    docwen_final = docwen.get("final") if isinstance(docwen, dict) else None
    if not isinstance(docwen_final, dict):
        raise V4PackageInputError("checkpoint_final_identities_missing")
    docwen_commit = str(docwen_final.get("commit", ""))
    docwen_tree = str(docwen_final.get("tree", ""))
    contract.validate_candidate_id(candidate_id, docwen_commit=docwen_commit)
    if any(marker in candidate_id.casefold() for marker in ("attempt04", "c6v", "c6-v")):
        raise V4PackageInputError("legacy_candidate_id_rejected")
    preparing = output.parent / f".{output.name}.prepare-{uuid.uuid4().hex}"
    work.mkdir(parents=False)
    published = False
    try:
        docwen_clone = work / "docwen-source"
        clone_factory(repositories[0], docwen_clone, docwen_commit, docwen_tree, "docwen")
        clone_verifier(docwen_clone, docwen_commit, docwen_tree, "docwen_post_clone")
        build = package_builder(docwen_clone, python, uv)
        clone_verifier(docwen_clone, docwen_commit, docwen_tree, "docwen_post_build")
        preparing.mkdir()
        package_root = preparing / "package-input"
        package_root.mkdir()
        _copy_tree_stable(build.gui, package_root / PACKAGE_NAMES[0])
        _copy_tree_stable(build.cli, package_root / PACKAGE_NAMES[1])
        standalone_cli = package_root / PACKAGE_NAMES[1] / "DocWenCLI.exe"
        bundled_cli = package_root / PACKAGE_NAMES[0] / "DocWenCLI.exe"
        gui_binary = package_root / PACKAGE_NAMES[0] / "DocWen.exe"
        for binary in (standalone_cli, bundled_cli, gui_binary):
            if not binary.is_file():
                raise V4PackageInputError(f"package_binary_missing:{binary.name}")
        standalone_identity = _identity(standalone_cli, root=package_root)
        bundled_identity = _identity(bundled_cli, root=package_root)
        gui_identity = _identity(gui_binary, root=package_root)
        if (
            standalone_identity["bytes"] != bundled_identity["bytes"]
            or standalone_identity["sha256"] != bundled_identity["sha256"]
        ):
            raise V4PackageInputError("standalone_and_bundled_cli_bytes_differ")
        package_tree = evidence_contract.capture_tree_stable(package_root)
        standalone_version = version_reader(standalone_cli)
        if _identity(standalone_cli, root=package_root) != standalone_identity:
            raise V4PackageInputError("standalone_cli_changed_during_version_probe")
        bundled_version = version_reader(bundled_cli)
        if _identity(bundled_cli, root=package_root) != bundled_identity:
            raise V4PackageInputError("bundled_cli_changed_during_version_probe")
        if _identity(gui_binary, root=package_root) != gui_identity:
            raise V4PackageInputError("gui_binary_changed_during_version_probe")
        if evidence_contract.capture_tree_stable(package_root) != package_tree:
            raise V4PackageInputError("package_changed_during_version_probe")
        if standalone_version != bundled_version:
            raise V4PackageInputError("standalone_and_bundled_cli_versions_differ")
        harness = _load_harness(docwen_clone)
        _revalidate_harness(harness, docwen=docwen_clone)
        harness_output = harness_runner(standalone_cli, docwen_clone, work / "harness-run", harness)
        _revalidate_harness(harness, docwen=docwen_clone)
        clone_verifier(docwen_clone, docwen_commit, docwen_tree, "docwen_post_harness")
        if evidence_contract.capture_tree_stable(package_root) != package_tree:
            raise V4PackageInputError("package_changed_during_harness")
        if _identity(standalone_cli, root=package_root) != standalone_identity:
            raise V4PackageInputError("standalone_cli_changed_during_harness")
        evidence_root = preparing / "evidence-input"
        evidence_root.mkdir()
        plan, package_manifest, record_summary = _build_evidence(
            evidence_root=evidence_root,
            package_root=package_root,
            docwen_clone=docwen_clone,
            checkpoint=checkpoint,
            candidate_id=candidate_id,
            harness=harness,
            harness_output=harness_output,
            executable_identity=standalone_identity,
            package_names=PACKAGE_NAMES,
            plan_schema=PLAN_SCHEMA,
        )
        _revalidate_harness(harness, docwen=docwen_clone)
        clone_verifier(docwen_clone, docwen_commit, docwen_tree, "docwen_post_evidence")
        if evidence_contract.capture_tree_stable(package_root) != package_tree:
            raise V4PackageInputError("package_changed_during_evidence")
        if (
            candidate_evidence.capture_package_manifest(
                package_root,
                PACKAGE_NAMES,
                allowed_root_entries=(),
            )
            != package_manifest
        ):
            raise V4PackageInputError("package_manifest_changed_during_evidence")
        plan_path = preparing / "evidence-plan.json"
        package_manifest_path = preparing / "package-manifest.preview.json"
        _write_json(plan_path, plan)
        _write_json(package_manifest_path, package_manifest)
        plan_identity = _identity(plan_path, root=preparing)
        package_manifest_identity = _identity(package_manifest_path, root=preparing)
        result = {
            "schema": PRODUCER_RESULT_SCHEMA,
            "candidateId": candidate_id,
            "productVersion": contract.PRODUCT_VERSION,
            "harness": {
                "id": HARNESS_ID,
                "version": HARNESS_VERSION,
                "schema": HARNESS_SCHEMA,
                "manifest": harness.manifest_identity,
                "requestSha256": harness_output.request_digest,
                "transcript": {
                    "bytes": len(harness_output.transcript),
                    "sha256": _sha256_bytes(harness_output.transcript),
                },
                "requiredCaseIds": list(REQUIRED_HARNESS_CASE_IDS),
                "executedCaseIds": [case.case_id for case in harness_output.cases],
                "pendingCaseIds": [],
                "forwardInputRoles": ["neutral_document", "numbering_export_plan"],
                "legacyForwardRolesRejected": ["source", "bibliography"],
            },
            "authority": {
                "sourceCheckpoint": checkpoint_identity,
                "docwen": {"commit": docwen_commit, "tree": docwen_tree},
                "finalSpecIdentity": checkpoint.get("finalSpecIdentity"),
            },
            "build": build.metadata,
            "package": {
                "root": "package-input",
                "manifestPreview": package_manifest_identity,
                "gui": gui_identity,
                "standaloneCli": {**standalone_identity, "version": standalone_version},
                "bundledCli": {**bundled_identity, "version": bundled_version},
                "cliBytesIdentical": True,
            },
            "evidence": {
                "root": "evidence-input",
                "plan": plan_identity,
                "records": record_summary,
                "hostStatus": dict.fromkeys(evidence_contract.HOST_LAYERS, "not_run"),
            },
        }
        _reject_legacy_harness({"harness": result["harness"], "plan": plan})
        result_path = preparing / "producer-result.json"
        _write_json(result_path, result)
        final_checkpoint, final_checkpoint_identity = checkpoint_loader(
            source_checkpoint,
            source_checkpoint_sha256,
            repositories[0],
        )
        if final_checkpoint != checkpoint or final_checkpoint_identity != checkpoint_identity:
            raise V4PackageInputError("source_checkpoint_changed_before_publish")
        _revalidate_harness(harness, docwen=docwen_clone)
        clone_verifier(docwen_clone, docwen_commit, docwen_tree, "docwen_pre_publish")
        if evidence_contract.capture_tree_stable(package_root) != package_tree:
            raise V4PackageInputError("package_changed_before_publish")
        stable = evidence_contract.capture_tree_stable(preparing)
        preparing.rename(output)
        published = True
        if evidence_contract.capture_tree_stable(output) != stable:
            raise V4PackageInputError("producer_output_changed_after_publish")
        published_checkpoint, published_checkpoint_identity = checkpoint_loader(
            source_checkpoint,
            source_checkpoint_sha256,
            repositories[0],
        )
        if published_checkpoint != checkpoint or published_checkpoint_identity != checkpoint_identity:
            raise V4PackageInputError("source_checkpoint_changed_after_publish")
        _revalidate_harness(harness, docwen=docwen_clone)
        clone_verifier(docwen_clone, docwen_commit, docwen_tree, "docwen_post_publish")
        published_package = output / "package-input"
        if evidence_contract.capture_tree_stable(published_package) != package_tree:
            raise V4PackageInputError("package_changed_after_publish")
        if (
            candidate_evidence.capture_package_manifest(
                published_package,
                PACKAGE_NAMES,
                allowed_root_entries=(),
            )
            != package_manifest
        ):
            raise V4PackageInputError("package_manifest_changed_after_publish")
        return {
            "candidateId": candidate_id,
            "outputRoot": str(output),
            "packageInput": str(output / "package-input"),
            "evidenceInput": str(output / "evidence-input"),
            "evidencePlan": {
                **_identity(output / "evidence-plan.json", root=output),
                "absolutePath": str(output / "evidence-plan.json"),
            },
            "producerResult": _identity(output / "producer-result.json", root=output),
            "treeManifestSha256": stable["manifestSha256"],
        }
    except BaseException:
        if preparing.exists() or preparing.is_symlink():
            rejected = output.parent / f".{output.name}.rejected-{uuid.uuid4().hex}"
            preparing.rename(rejected)
        elif published and (output.exists() or output.is_symlink()):
            rejected = output.parent / f".{output.name}.rejected-{uuid.uuid4().hex}"
            output.rename(rejected)
        raise


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--docwen-repo", type=Path, required=True)
    parser.add_argument("--source-checkpoint", type=Path, required=True)
    parser.add_argument("--source-checkpoint-sha256", required=True)
    parser.add_argument("--candidate-id", required=True)
    parser.add_argument("--python", type=Path, required=True)
    parser.add_argument("--uv", type=Path, required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    parser.add_argument("--work-root", type=Path, required=True)
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        result = build_package_input(
            docwen_repo=args.docwen_repo,
            source_checkpoint=args.source_checkpoint,
            source_checkpoint_sha256=args.source_checkpoint_sha256,
            candidate_id=args.candidate_id,
            python=args.python,
            uv=args.uv,
            output_root=args.output_root,
            work_root=args.work_root,
        )
    except (
        V4PackageInputError,
        contract.V4CandidateContractError,
        evidence_contract.V4EvidenceContractError,
        candidate_evidence.EvidenceError,
        OSError,
        subprocess.SubprocessError,
    ) as exc:
        print(f"build_v4_package_input_error:{exc}", file=sys.stderr)
        return 2
    print(json.dumps({"success": True, "data": result}, ensure_ascii=False, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
