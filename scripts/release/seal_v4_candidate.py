from __future__ import annotations

import argparse
import json
import shutil
import sys
import uuid
from collections.abc import Mapping
from datetime import UTC, datetime
from pathlib import Path
from typing import Any, cast

if __package__ in {None, ""}:
    _BOOTSTRAP_ROOT = Path(__file__).resolve().parents[2]
    if str(_BOOTSTRAP_ROOT) not in sys.path:
        sys.path.insert(0, str(_BOOTSTRAP_ROOT))

from scripts.release import candidate_evidence
from scripts.release import v4_candidate_contract as contract

SOURCE_CHECKPOINT_SCHEMA = "docwen.v4_candidate_source_checkpoint.v1"
STAGING_CHECKPOINT_SCHEMA = "docwen.v4_candidate_staging_checkpoint.v1"
CANDIDATE_SCHEMA = "docwen.local_candidate.v4"
OUTER_MANIFEST_SCHEMA = "docwen.v4_candidate_outer_manifest.v1"


class V4CandidateSealError(RuntimeError):
    """A v4 candidate could not be sealed without mutating package bytes."""


def _safe_new_parent(path: Path, *, label: str) -> Path:
    if not path.name or ":" in path.name or "\\" in path.name:
        raise V4CandidateSealError(f"{label}_path_invalid:{path}")
    parent = candidate_evidence._safe_existing_directory(path.parent, label=f"{label}_parent")
    candidate = parent / path.name
    if candidate.exists() or candidate.is_symlink():
        raise V4CandidateSealError(f"{label}_already_exists:{candidate}")
    return candidate


def _copy_stable(source: Path, destination: Path) -> dict[str, object]:
    source_tree = contract.capture_tree_stable(source)
    shutil.copytree(source, destination, symlinks=True)
    if contract.capture_tree_stable(source) != source_tree:
        raise V4CandidateSealError("source_staging_changed_during_copy")
    if contract.capture_tree_stable(destination) != source_tree:
        raise V4CandidateSealError("staging_copy_bytes_mismatch")
    return source_tree


def _source_checkpoint_payload(docwen_repo: Path) -> dict[str, object]:
    return contract.capture_source_checkpoint(docwen_repo=docwen_repo)


def write_source_checkpoint(*, docwen_repo: Path, output: Path) -> dict[str, object]:
    safe_output = _safe_new_parent(output, label="source_checkpoint")
    protected = docwen_repo.resolve(strict=True)
    if safe_output == protected or protected in safe_output.parents:
        raise V4CandidateSealError("source_checkpoint_must_be_outside_source_repository")
    payload = _source_checkpoint_payload(docwen_repo)
    contract.write_json_exclusive(safe_output, payload)
    return {"checkpoint": contract.file_identity(safe_output, relative_to=safe_output.parent)}


def _read_external_checkpoint(path: Path, *, expected_sha256: str, label: str) -> tuple[dict[str, Any], dict[str, Any]]:
    contract.require_hex64(expected_sha256, label=label)
    try:
        value, identity = contract.read_json_object_with_identity(
            path,
            relative_to=path.parent,
            label=label,
            expected_sha256=expected_sha256,
        )
        return value, cast(dict[str, Any], identity)
    except contract.V4CandidateContractError as exc:
        raise V4CandidateSealError(str(exc)) from exc


def _package_manifest(
    staging_root: Path, *, stable_tree: Mapping[str, object] | None = None
) -> tuple[dict[str, Any], dict[str, str | int]]:
    manifest_path = staging_root / "_evidence" / "package-manifest.json"
    before = stable_tree or contract.capture_tree_stable(staging_root)
    if stable_tree is not None and contract.capture_tree_stable(staging_root) != stable_tree:
        raise V4CandidateSealError("staging_tree_changed_before_package_verification")
    allowed = tuple(
        sorted(
            item.name
            for item in staging_root.iterdir()
            if item.name not in {"DocWen_v0.9.0_win-x64", "DocWenCLI_v0.9.0_win-x64"}
            and item.is_dir()
            and not item.is_symlink()
        )
    )
    manifest = contract.read_json_object(manifest_path, label="package_manifest")
    candidate_evidence.verify_package_manifest(staging_root, manifest_path, allowed_root_entries=allowed)
    after = contract.capture_tree_stable(staging_root)
    if before != after:
        raise V4CandidateSealError("staging_tree_changed_during_package_verification")
    return manifest, contract.identity_from_tree(manifest_path, relative_to=staging_root, stable_tree=after)


def _payload_tree(staging_root: Path) -> dict[str, object]:
    raw = contract.capture_tree_stable(staging_root)
    raw_files = raw.get("files")
    if not isinstance(raw_files, list):
        raise V4CandidateSealError("staging_tree_files_invalid")
    excluded = {"candidate.json", "manifest.json", "receipt.json"}
    files = [
        item
        for item in raw_files
        if isinstance(item, dict)
        and not any(
            str(item.get("path")) == prefix or str(item.get("path")).startswith(f"{prefix}/") for prefix in excluded
        )
    ]
    return {
        "schema": raw.get("schema"),
        "fileCount": len(files),
        "totalBytes": sum(cast(int, item["size"]) for item in files),
        "manifestSha256": candidate_evidence._payload_hash(files),
        "files": files,
    }


def write_staging_checkpoint(
    *,
    docwen_repo: Path,
    staging_root: Path,
    source_checkpoint: Path,
    source_checkpoint_sha256: str,
    output: Path,
) -> dict[str, object]:
    safe_output = _safe_new_parent(output, label="staging_checkpoint")
    safe_staging = staging_root.resolve(strict=True)
    if safe_output == safe_staging or safe_staging in safe_output.parents:
        raise V4CandidateSealError("staging_checkpoint_must_be_outside_staging")
    source, source_identity = _read_external_checkpoint(
        source_checkpoint, expected_sha256=source_checkpoint_sha256, label="source_checkpoint"
    )
    contract.verify_source_checkpoint(source, docwen_repo=docwen_repo)
    if source.get("schema") != SOURCE_CHECKPOINT_SCHEMA:
        raise V4CandidateSealError("source_checkpoint_schema_mismatch")
    manifest, manifest_identity = _package_manifest(safe_staging)
    staging_tree = _payload_tree(safe_staging)
    payload = {
        "schema": STAGING_CHECKPOINT_SCHEMA,
        "stagingRootName": safe_staging.name,
        "sourceCheckpoint": source_identity,
        "sourceCheckpointPayloadSha256": contract.payload_sha256(source),
        "sourceAuthority": {
            "docwen": source["docwen"],
            "activeSemantics": source["activeSemantics"],
        },
        "packageManifest": manifest_identity,
        "packageManifestPayloadSha256": contract.payload_sha256(manifest),
        "stagingTree": staging_tree,
    }
    contract.write_json_exclusive(safe_output, payload)
    return {"checkpoint": contract.file_identity(safe_output, relative_to=safe_output.parent)}


def _verify_staged_governance(staging_root: Path, source: Mapping[str, object]) -> None:
    final_spec = source.get("finalSpecIdentity")
    raw_spec_files = final_spec.get("files") if isinstance(final_spec, dict) else None
    if not isinstance(raw_spec_files, list):
        raise V4CandidateSealError("source_final_spec_identity_invalid")
    identities = [source.get("authorityBaseline"), source.get("oracleManifest"), *raw_spec_files]
    expected: dict[str, Mapping[str, object]] = {}
    for raw in identities:
        if not isinstance(raw, dict) or not isinstance(raw.get("relativePath"), str):
            raise V4CandidateSealError("source_governance_identity_invalid")
        expected[cast(str, raw["relativePath"])] = cast(Mapping[str, object], raw)
    for required in (contract.RECEIPT_SCHEMA_PATH, contract.EVIDENCE_SCHEMA_PATH):
        if required not in expected:
            raise V4CandidateSealError(f"source_candidate_schema_identity_missing:{required}")
    for relative, identity in expected.items():
        actual = contract.file_identity(staging_root / relative, relative_to=staging_root)
        if actual != identity:
            raise V4CandidateSealError(f"staged_governance_bytes_mismatch:{relative}")


def _verify_staging_checkpoint(
    checkpoint: Mapping[str, object],
    *,
    staging_root: Path,
    original_staging_name: str,
    source: Mapping[str, object],
    source_identity: Mapping[str, object],
    stable_tree: Mapping[str, object] | None = None,
) -> tuple[dict[str, Any], dict[str, str | int]]:
    if (
        checkpoint.get("schema") != STAGING_CHECKPOINT_SCHEMA
        or checkpoint.get("stagingRootName") != original_staging_name
    ):
        raise V4CandidateSealError("staging_checkpoint_identity_mismatch")
    if checkpoint.get("sourceCheckpoint") != source_identity:
        raise V4CandidateSealError("staging_checkpoint_source_file_mismatch")
    if checkpoint.get("sourceCheckpointPayloadSha256") != contract.payload_sha256(source):
        raise V4CandidateSealError("staging_checkpoint_source_payload_mismatch")
    _verify_staged_governance(staging_root, source)
    manifest, manifest_identity = _package_manifest(staging_root, stable_tree=stable_tree)
    if checkpoint.get("packageManifest") != manifest_identity:
        raise V4CandidateSealError("staging_checkpoint_package_manifest_mismatch")
    if checkpoint.get("packageManifestPayloadSha256") != contract.payload_sha256(manifest):
        raise V4CandidateSealError("staging_checkpoint_package_payload_mismatch")
    if checkpoint.get("stagingTree") != _payload_tree(staging_root):
        raise V4CandidateSealError("staging_checkpoint_tree_mismatch")
    return manifest, manifest_identity


def _find_cli_identities(
    candidate_root: Path, *, stable_tree: Mapping[str, object] | None = None
) -> tuple[dict[str, str | int], dict[str, str | int]]:
    standalone = list(candidate_root.glob("DocWenCLI_v*_win-x64/DocWenCLI.exe"))
    bundled = list(candidate_root.glob("DocWen_v*_win-x64/DocWenCLI.exe"))
    if len(standalone) != 1 or len(bundled) != 1:
        raise V4CandidateSealError("candidate_cli_cardinality_invalid")
    tree = stable_tree or contract.capture_tree_stable(candidate_root)
    standalone_identity = contract.identity_from_tree(standalone[0], relative_to=candidate_root, stable_tree=tree)
    bundled_identity = contract.identity_from_tree(bundled[0], relative_to=candidate_root, stable_tree=tree)
    if (
        standalone_identity["bytes"] != bundled_identity["bytes"]
        or standalone_identity["sha256"] != bundled_identity["sha256"]
    ):
        raise V4CandidateSealError("candidate_cli_bytes_differ")
    return standalone_identity, bundled_identity


def _evidence_index(
    candidate_root: Path, *, candidate_id: str, stable_tree: Mapping[str, object]
) -> tuple[dict[str, Any], dict[str, str | int]]:
    path = candidate_root / "_evidence" / "v4-evidence-index.json"
    value = contract.read_json_object(path, label="v4_evidence_index")
    contract.validate_json_schema(value, candidate_root / contract.EVIDENCE_SCHEMA_PATH)
    records, _ = contract.validate_evidence_index(value, candidate_id=candidate_id)
    contract.verify_index_files(candidate_root, records, stable_tree=stable_tree)
    return value, contract.identity_from_tree(path, relative_to=candidate_root, stable_tree=stable_tree)


def _receipt_payload(
    *,
    candidate_id: str,
    candidate_root: Path,
    source: Mapping[str, object],
    manifest_identity: Mapping[str, object],
    evidence_index: Mapping[str, object],
    evidence_identity: Mapping[str, object],
    stable_tree: Mapping[str, object],
) -> dict[str, object]:
    standalone_cli, bundled_cli = _find_cli_identities(candidate_root, stable_tree=stable_tree)
    records, invalid_records = contract.validate_evidence_index(evidence_index, candidate_id=candidate_id)
    layer_status = cast(Mapping[str, object], evidence_index["layerStatus"])
    eligible = contract.consumer_eligible(layer_status)
    return {
        "schema": contract.RECEIPT_SCHEMA,
        "candidateId": candidate_id,
        "state": "local_unpublished_complete" if eligible else "local_unpublished_host_validation_pending",
        "consumerEligible": eligible,
        "productVersion": contract.PRODUCT_VERSION,
        "authority": {
            "activeSemantics": source["activeSemantics"],
            "authorityBaseline": source["authorityBaseline"],
            "finalSpecIdentity": source["finalSpecIdentity"],
            "docwen": source["docwen"],
            "exclusions": source["exclusions"],
        },
        "package": {
            "manifest": dict(manifest_identity),
            "standaloneCli": standalone_cli,
            "bundledCli": bundled_cli,
            "cliBytesIdentical": True,
        },
        "machineConsumerContract": contract.machine_consumer_contract(),
        "evidence": {
            "index": dict(evidence_identity),
            "layerStatus": dict(layer_status),
            "records": records,
            "sixIndependentInvalidIdRecords": invalid_records,
        },
        "hashDag": {
            "acyclic": True,
            "receiptOmitsPostReceiptRecords": ["candidate", "outer_manifest"],
            "candidateBindsReceipt": True,
            "outerManifestBindsBoth": True,
        },
    }


def _outer_manifest_payload(candidate_root: Path, *, candidate_id: str) -> dict[str, object]:
    tree = contract.capture_tree_stable(candidate_root)
    receipt = contract.identity_from_tree(candidate_root / "receipt.json", relative_to=candidate_root, stable_tree=tree)
    candidate = contract.identity_from_tree(
        candidate_root / "candidate.json", relative_to=candidate_root, stable_tree=tree
    )
    package = contract.identity_from_tree(
        candidate_root / "_evidence" / "package-manifest.json", relative_to=candidate_root, stable_tree=tree
    )
    cli, _ = _find_cli_identities(candidate_root, stable_tree=tree)
    return {
        "schema": OUTER_MANIFEST_SCHEMA,
        "candidateId": candidate_id,
        "productVersion": contract.PRODUCT_VERSION,
        "receipt": receipt,
        "candidate": candidate,
        "packageManifest": package,
        "standaloneCli": cli,
    }


def _atomic_publish(*, publishing_root: Path, final_root: Path, expected_tree: Mapping[str, object]) -> None:
    if final_root.exists() or final_root.is_symlink():
        raise V4CandidateSealError("final_candidate_path_exists")
    if contract.capture_tree_stable(publishing_root) != expected_tree:
        raise V4CandidateSealError("publishing_tree_changed_before_publish")
    publishing_root.rename(final_root)
    try:
        if contract.capture_tree_stable(final_root) != expected_tree:
            raise V4CandidateSealError("candidate_tree_changed_after_publish")
    except BaseException as proof_error:
        quarantine = final_root.parent / f".rejected-{final_root.name}-{uuid.uuid4().hex}"
        try:
            final_root.rename(quarantine)
        except BaseException as quarantine_error:
            error = V4CandidateSealError(
                f"post_publish_proof_failed_and_quarantine_failed:candidate={final_root}:quarantine={quarantine}"
            )
            error.add_note(f"post-publish proof: {proof_error}")
            raise error from quarantine_error
        raise V4CandidateSealError(f"candidate_quarantined_after_failed_proof:{quarantine}") from proof_error


def seal_candidate(
    *,
    docwen_repo: Path,
    staging_root: Path,
    source_checkpoint: Path,
    source_checkpoint_sha256: str,
    staging_checkpoint: Path,
    staging_checkpoint_sha256: str,
    output_root: Path,
    candidate_id: str,
    generated_at: str,
) -> dict[str, object]:
    output_parent = output_root.resolve(strict=True)
    source_staging = staging_root.resolve(strict=True)
    if (
        source_staging == output_parent
        or source_staging in output_parent.parents
        or output_parent in source_staging.parents
    ):
        raise V4CandidateSealError("output_and_staging_roots_overlap")
    source, source_identity = _read_external_checkpoint(
        source_checkpoint, expected_sha256=source_checkpoint_sha256, label="source_checkpoint"
    )
    checkpoint, _ = _read_external_checkpoint(
        staging_checkpoint, expected_sha256=staging_checkpoint_sha256, label="staging_checkpoint"
    )
    contract.verify_source_checkpoint(source, docwen_repo=docwen_repo)
    docwen = cast(Mapping[str, object], source["docwen"])
    final_docwen = cast(Mapping[str, object], docwen["final"])
    contract.validate_candidate_id(candidate_id, docwen_commit=str(final_docwen["commit"]))
    _verify_staging_checkpoint(
        checkpoint,
        staging_root=source_staging,
        original_staging_name=source_staging.name,
        source=source,
        source_identity=source_identity,
    )
    publishing = _safe_new_parent(output_parent / f".publishing-{candidate_id}", label="publishing_root")
    final_root = _safe_new_parent(output_parent / candidate_id, label="candidate_root")
    try:
        _copy_stable(source_staging, publishing)
        _, manifest_identity = _verify_staging_checkpoint(
            checkpoint,
            staging_root=publishing,
            original_staging_name=source_staging.name,
            source=source,
            source_identity=source_identity,
        )
        sealed_input_tree = contract.capture_tree_stable(publishing)
        _, manifest_identity = _verify_staging_checkpoint(
            checkpoint,
            staging_root=publishing,
            original_staging_name=source_staging.name,
            source=source,
            source_identity=source_identity,
            stable_tree=sealed_input_tree,
        )
        evidence_index, evidence_identity = _evidence_index(
            publishing, candidate_id=candidate_id, stable_tree=sealed_input_tree
        )
        receipt = _receipt_payload(
            candidate_id=candidate_id,
            candidate_root=publishing,
            source=source,
            manifest_identity=manifest_identity,
            evidence_index=evidence_index,
            evidence_identity=evidence_identity,
            stable_tree=sealed_input_tree,
        )
        contract.validate_json_schema(receipt, publishing / contract.RECEIPT_SCHEMA_PATH)
        receipt_path = publishing / "receipt.json"
        contract.write_json_exclusive(receipt_path, receipt)
        receipt_identity = contract.file_identity(receipt_path, relative_to=publishing)
        candidate_payload = {
            "schema": CANDIDATE_SCHEMA,
            "candidateId": candidate_id,
            "generatedAt": generated_at,
            "state": receipt["state"],
            "consumerEligible": receipt["consumerEligible"],
            "productVersion": contract.PRODUCT_VERSION,
            "receipt": receipt_identity,
        }
        contract.write_json_exclusive(publishing / "candidate.json", candidate_payload)
        outer = _outer_manifest_payload(publishing, candidate_id=candidate_id)
        contract.write_json_exclusive(publishing / "manifest.json", outer)
        if contract.capture_tree_stable(publishing) == sealed_input_tree:
            raise V4CandidateSealError("post_receipt_records_missing_from_candidate_tree")
        if contract.read_json_object(receipt_path, label="final_receipt") != receipt:
            raise V4CandidateSealError("receipt_changed_after_outer_records")
        expected_tree = contract.capture_tree_stable(publishing)
        contract.verify_source_checkpoint(source, docwen_repo=docwen_repo)
        if _payload_tree(source_staging) != checkpoint["stagingTree"]:
            raise V4CandidateSealError("source_staging_changed_before_publish")
        _atomic_publish(publishing_root=publishing, final_root=final_root, expected_tree=expected_tree)
    except BaseException as proof_error:
        if publishing.exists() or publishing.is_symlink():
            rejected = output_parent / f".rejected-{candidate_id}-{uuid.uuid4().hex}"
            try:
                publishing.rename(rejected)
            except BaseException as quarantine_error:
                error = V4CandidateSealError("seal_failure_and_quarantine_failed:canonical_absent_or_nonreusable")
                error.add_note(f"seal failure: {proof_error}")
                error.add_note(f"quarantine failure: {quarantine_error}")
                raise error from quarantine_error
        raise
    final_tree = contract.capture_tree_stable(final_root)
    return {
        "candidateId": candidate_id,
        "candidateRoot": str(final_root),
        "consumerEligible": receipt["consumerEligible"],
        "productVersion": contract.PRODUCT_VERSION,
        "docWenCli": contract.file_identity(
            next(final_root.glob("DocWenCLI_v*_win-x64/DocWenCLI.exe")), relative_to=final_root
        ),
        "candidate": contract.file_identity(final_root / "candidate.json", relative_to=final_root),
        "manifest": contract.file_identity(final_root / "manifest.json", relative_to=final_root),
        "receipt": contract.file_identity(final_root / "receipt.json", relative_to=final_root),
        "treeManifestSha256": final_tree["manifestSha256"],
    }


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Checkpoint or seal one existing, already-tested DocWen v4 staging tree."
    )
    subparsers = parser.add_subparsers(dest="command", required=True)
    source = subparsers.add_parser("source-checkpoint")
    source.add_argument("--docwen-repo", type=Path, required=True)
    source.add_argument("--output", type=Path, required=True)
    staging = subparsers.add_parser("staging-checkpoint")
    staging.add_argument("--docwen-repo", type=Path, required=True)
    staging.add_argument("--staging-root", type=Path, required=True)
    staging.add_argument("--source-checkpoint", type=Path, required=True)
    staging.add_argument("--source-checkpoint-sha256", required=True)
    staging.add_argument("--output", type=Path, required=True)
    seal = subparsers.add_parser("seal")
    seal.add_argument("--docwen-repo", type=Path, required=True)
    seal.add_argument("--staging-root", type=Path, required=True)
    seal.add_argument("--source-checkpoint", type=Path, required=True)
    seal.add_argument("--source-checkpoint-sha256", required=True)
    seal.add_argument("--staging-checkpoint", type=Path, required=True)
    seal.add_argument("--staging-checkpoint-sha256", required=True)
    seal.add_argument("--output-root", type=Path, required=True)
    seal.add_argument("--candidate-id", required=True)
    seal.add_argument("--generated-at", default=datetime.now(UTC).isoformat().replace("+00:00", "Z"))
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        if args.command == "source-checkpoint":
            result = write_source_checkpoint(
                docwen_repo=args.docwen_repo,
                output=args.output,
            )
        elif args.command == "staging-checkpoint":
            result = write_staging_checkpoint(
                docwen_repo=args.docwen_repo,
                staging_root=args.staging_root,
                source_checkpoint=args.source_checkpoint,
                source_checkpoint_sha256=args.source_checkpoint_sha256,
                output=args.output,
            )
        else:
            result = seal_candidate(
                docwen_repo=args.docwen_repo,
                staging_root=args.staging_root,
                source_checkpoint=args.source_checkpoint,
                source_checkpoint_sha256=args.source_checkpoint_sha256,
                staging_checkpoint=args.staging_checkpoint,
                staging_checkpoint_sha256=args.staging_checkpoint_sha256,
                output_root=args.output_root,
                candidate_id=args.candidate_id,
                generated_at=args.generated_at,
            )
    except (
        V4CandidateSealError,
        contract.V4CandidateContractError,
        candidate_evidence.EvidenceError,
        OSError,
    ) as exc:
        print(f"seal_v4_candidate_error:{exc}", file=sys.stderr)
        return 2
    print(json.dumps({"success": True, "data": result}, ensure_ascii=False, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
