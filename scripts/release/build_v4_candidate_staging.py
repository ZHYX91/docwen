from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import sys
import uuid
from collections.abc import Mapping
from pathlib import Path
from typing import Any, cast

if __package__ in {None, ""}:
    _BOOTSTRAP_ROOT = Path(__file__).resolve().parents[2]
    if str(_BOOTSTRAP_ROOT) not in sys.path:
        sys.path.insert(0, str(_BOOTSTRAP_ROOT))

from scripts.release import candidate_evidence
from scripts.release import seal_v4_candidate as sealer
from scripts.release import v4_candidate_contract as contract
from scripts.release import v4_evidence_contract as evidence_contract

PLAN_SCHEMA = "docwen.v4_evidence_plan.v1"
RECORD_SCHEMA = evidence_contract.RECORD_SCHEMA
EXTERNAL_MANIFEST_LAYERS = evidence_contract.EXTERNAL_MANIFEST_LAYERS
PACKAGE_NAMES = (
    f"DocWen_v{contract.PRODUCT_VERSION}_win-x64",
    f"DocWenCLI_v{contract.PRODUCT_VERSION}_win-x64",
)


class V4StagingBuildError(RuntimeError):
    """A v4 evidence staging tree could not be produced without ambiguity."""


def _safe_relative(value: object, *, label: str) -> str:
    try:
        return evidence_contract.validate_relative_path(value, label=label)
    except evidence_contract.V4EvidenceContractError as exc:
        raise V4StagingBuildError(str(exc)) from exc


def _read_external(path: Path, *, expected_sha256: str, label: str) -> tuple[dict[str, Any], dict[str, object]]:
    contract.require_hex64(expected_sha256, label=label)
    try:
        return contract.read_json_object_with_identity(
            path,
            relative_to=path.parent,
            label=label,
            expected_sha256=expected_sha256,
        )
    except contract.V4CandidateContractError as exc:
        raise V4StagingBuildError(str(exc)) from exc


_record_destination = evidence_contract.record_path
_manifest_destination = evidence_contract.manifest_path


def _json_identity(value: object, *, relative_path: str) -> dict[str, object]:
    raw = contract.json_bytes(value)
    return {"relativePath": relative_path, "bytes": len(raw), "sha256": hashlib.sha256(raw).hexdigest()}


def _identity_at(identity: Mapping[str, object], relative_path: str) -> dict[str, object]:
    return {"relativePath": relative_path, "bytes": identity["bytes"], "sha256": identity["sha256"]}


_with_role = evidence_contract.pointer_identity


def _validate_record_envelope(value: Mapping[str, object], *, case_id: str, layer: str) -> None:
    if not evidence_contract.record_envelope_matches(value, case_id=case_id, layer=layer):
        raise V4StagingBuildError(f"evidence_record_envelope_mismatch:{layer}:{case_id}")


def _parse_plan(
    value: Mapping[str, object], *, candidate_id: str
) -> tuple[dict[str, str], list[dict[str, str]], dict[str, str | None]]:
    if set(value) != {"schema", "candidateId", "layerStatus", "records", "manifests"}:
        raise V4StagingBuildError("evidence_plan_not_closed")
    if value.get("schema") != PLAN_SCHEMA or value.get("candidateId") != candidate_id:
        raise V4StagingBuildError("evidence_plan_identity_mismatch")
    statuses = value.get("layerStatus")
    raw_records = value.get("records")
    manifests = value.get("manifests")
    if not isinstance(statuses, dict) or set(statuses) != set(contract.REQUIRED_LAYERS):
        raise V4StagingBuildError("evidence_plan_layer_status_invalid")
    if any(statuses.get(layer) != "passed" for layer in contract.REQUIRED_LAYERS[:6]):
        raise V4StagingBuildError("evidence_plan_non_host_layer_not_passed")
    if any(statuses.get(layer) not in {"passed", "not_run"} for layer in contract.REQUIRED_LAYERS[6:]):
        raise V4StagingBuildError("evidence_plan_host_status_invalid")
    if not isinstance(raw_records, list) or not raw_records or not isinstance(manifests, dict):
        raise V4StagingBuildError("evidence_plan_records_or_manifests_invalid")
    if set(manifests) != set(EXTERNAL_MANIFEST_LAYERS):
        raise V4StagingBuildError("evidence_plan_manifest_layers_invalid")
    records: list[dict[str, str]] = []
    keys: set[tuple[str, str]] = set()
    for raw in raw_records:
        if not isinstance(raw, dict) or set(raw) != {"caseId", "layer", "artifact"}:
            raise V4StagingBuildError("evidence_plan_record_not_closed")
        case_id = raw.get("caseId")
        layer = raw.get("layer")
        if not isinstance(case_id, str) or not isinstance(layer, str) or layer not in contract.REQUIRED_LAYERS:
            raise V4StagingBuildError("evidence_plan_record_identity_invalid")
        key = (layer, case_id)
        if key in keys:
            raise V4StagingBuildError(f"evidence_plan_duplicate_record:{layer}:{case_id}")
        keys.add(key)
        records.append(
            {"caseId": case_id, "layer": layer, "artifact": _safe_relative(raw.get("artifact"), label="artifact")}
        )
    typed_manifests: dict[str, str | None] = {}
    for layer in EXTERNAL_MANIFEST_LAYERS:
        raw_path = manifests[layer]
        status = cast(dict[str, str], statuses)[layer]
        if status == "not_run":
            if raw_path is not None:
                raise V4StagingBuildError(f"not_run_host_manifest_declared:{layer}")
            typed_manifests[layer] = None
        else:
            typed_manifests[layer] = _safe_relative(raw_path, label=f"{layer}_manifest")
    typed_statuses = cast(dict[str, str], statuses)
    for layer in contract.REQUIRED_LAYERS:
        matches = [item for item in records if item["layer"] == layer]
        expected = 1 if layer in contract.REQUIRED_LAYERS[6:] and typed_statuses[layer] == "passed" else None
        if layer in contract.REQUIRED_LAYERS[6:] and len(matches) != (expected or 0):
            raise V4StagingBuildError(f"host_record_cardinality_invalid:{layer}:{len(matches)}")
        if layer not in contract.REQUIRED_LAYERS[6:] and not matches:
            raise V4StagingBuildError(f"non_host_record_missing:{layer}")
    for dimension, case_id in contract.REQUIRED_INVALID_CASES.items():
        if ("source_oracle", case_id) not in keys:
            raise V4StagingBuildError(f"independent_invalid_record_missing:{dimension}")
    return typed_statuses, records, typed_manifests


def _governance_identities(source: Mapping[str, object]) -> dict[str, Mapping[str, object]]:
    final_spec = source.get("finalSpecIdentity")
    raw_files = final_spec.get("files") if isinstance(final_spec, dict) else None
    if not isinstance(raw_files, list):
        raise V4StagingBuildError("source_final_spec_files_invalid")
    raw_identities: list[object] = [source.get("authorityBaseline"), source.get("oracleManifest"), *raw_files]
    identities: dict[str, Mapping[str, object]] = {}
    for raw in raw_identities:
        if not isinstance(raw, dict):
            raise V4StagingBuildError("source_governance_identity_invalid")
        relative = _safe_relative(raw.get("relativePath"), label="governance")
        if relative in identities and identities[relative] != raw:
            raise V4StagingBuildError(f"source_governance_identity_conflict:{relative}")
        identities[relative] = raw
    for required in (
        contract.AUTHORITY_PATH,
        contract.RECEIPT_SCHEMA_PATH,
        contract.EVIDENCE_SCHEMA_PATH,
        contract.WIRE_SCHEMA_PATH,
        str(cast(Mapping[str, object], source["oracleManifest"])["relativePath"]),
    ):
        if required not in identities:
            raise V4StagingBuildError(f"source_governance_identity_missing:{required}")
    return identities


def _preflight_source_files(docwen_repo: Path, identities: Mapping[str, Mapping[str, object]]) -> None:
    for relative, expected in identities.items():
        if contract.file_identity(docwen_repo / relative, relative_to=docwen_repo) != expected:
            raise V4StagingBuildError(f"source_governance_bytes_drifted:{relative}")


def _copy_checked(
    source: Path,
    destination: Path,
    *,
    source_root: Path,
    destination_root: Path,
    expected: Mapping[str, object],
) -> dict[str, object]:
    if contract.file_identity(source, relative_to=source_root) != expected:
        raise V4StagingBuildError(f"copy_source_changed:{source}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists() or destination.is_symlink():
        actual = cast(dict[str, object], contract.file_identity(destination, relative_to=destination_root))
        expected_at_destination = _identity_at(expected, cast(str, actual["relativePath"]))
        if actual != expected_at_destination:
            raise V4StagingBuildError(f"copy_destination_conflict:{destination}")
        return actual
    shutil.copy2(source, destination)
    if contract.file_identity(source, relative_to=source_root) != expected:
        raise V4StagingBuildError(f"copy_source_changed_during_copy:{source}")
    actual = cast(dict[str, object], contract.file_identity(destination, relative_to=destination_root))
    if actual != _identity_at(expected, cast(str, actual["relativePath"])):
        raise V4StagingBuildError(f"copy_destination_bytes_mismatch:{destination}")
    return actual


_manifest_expected = evidence_contract.manifest_expected


def _quarantine(path: Path, *, output_root: Path, candidate_id: str) -> Path | None:
    if path.exists() or path.is_symlink():
        rejected = output_root / f".rejected-{candidate_id}-{uuid.uuid4().hex}"
        path.rename(rejected)
        return rejected
    return None


def build_staging(
    *,
    docwen_repo: Path,
    source_staging: Path,
    evidence_root: Path,
    evidence_plan: Path,
    evidence_plan_sha256: str,
    source_checkpoint: Path,
    source_checkpoint_sha256: str,
    staging_checkpoint_output: Path,
    output_root: Path,
    candidate_id: str,
) -> dict[str, object]:
    source, source_identity = _read_external(
        source_checkpoint, expected_sha256=source_checkpoint_sha256, label="source_checkpoint"
    )
    contract.verify_source_checkpoint(source, docwen_repo=docwen_repo)
    docwen = cast(Mapping[str, object], source["docwen"])
    final_docwen = cast(Mapping[str, object], docwen["final"])
    contract.validate_candidate_id(candidate_id, docwen_commit=str(final_docwen["commit"]))
    plan, plan_identity = _read_external(evidence_plan, expected_sha256=evidence_plan_sha256, label="evidence_plan")
    statuses, planned_records, planned_manifests = _parse_plan(plan, candidate_id=candidate_id)
    safe_evidence = candidate_evidence._safe_existing_directory(evidence_root, label="v4_evidence_root")
    safe_source_staging = candidate_evidence._safe_existing_directory(source_staging, label="source_staging")
    safe_output = candidate_evidence._safe_existing_directory(output_root, label="staging_output_root")
    source_repositories = (docwen_repo.resolve(strict=True),)
    for protected in (safe_source_staging, safe_evidence, *source_repositories):
        if protected == safe_output or protected in safe_output.parents or safe_output in protected.parents:
            raise V4StagingBuildError("staging_output_overlaps_input")
    for forbidden in ("evidence", "_evidence", "receipt.json", "candidate.json", "manifest.json"):
        if (safe_source_staging / forbidden).exists() or (safe_source_staging / forbidden).is_symlink():
            raise V4StagingBuildError(f"source_staging_contains_reserved_path:{forbidden}")
    final_root = safe_output / f"{candidate_id}-staging"
    preparing = safe_output / f".preparing-v4-{uuid.uuid4().hex[:12]}"
    if final_root.exists() or final_root.is_symlink():
        raise V4StagingBuildError("final_staging_path_exists")
    checkpoint_parent = staging_checkpoint_output.parent.resolve(strict=True)
    safe_checkpoint = checkpoint_parent / staging_checkpoint_output.name
    if safe_checkpoint.exists() or safe_checkpoint.is_symlink():
        raise V4StagingBuildError("staging_checkpoint_already_exists")
    checkpoint_preparing = safe_checkpoint.parent / f".preparing-{safe_checkpoint.name}-{uuid.uuid4().hex}"
    if checkpoint_preparing.exists() or checkpoint_preparing.is_symlink():  # pragma: no cover - UUID collision
        raise V4StagingBuildError("staging_checkpoint_preparing_path_exists")
    protected = (*source_repositories, safe_source_staging, safe_evidence, safe_output)
    if any(safe_checkpoint == root or root in safe_checkpoint.parents for root in protected):
        raise V4StagingBuildError("staging_checkpoint_not_external")
    safe_source_checkpoint = source_checkpoint.resolve(strict=True)
    if any(
        safe_source_checkpoint == root
        or root in safe_source_checkpoint.parents
        or safe_source_checkpoint in root.parents
        for root in protected
    ):
        raise V4StagingBuildError("source_checkpoint_not_external")
    source_tree = contract.capture_tree_stable(safe_source_staging)
    governance = _governance_identities(source)
    _preflight_source_files(docwen_repo, governance)
    evidence_sources: dict[tuple[str, str], tuple[Path, dict[str, object]]] = {}
    record_values: list[dict[str, Any]] = []
    predicted_records: list[dict[str, object]] = []
    for item in planned_records:
        source_path = safe_evidence / item["artifact"]
        value, identity = contract.read_json_object_with_identity(
            source_path,
            relative_to=safe_evidence,
            label=f"evidence_record:{item['caseId']}",
        )
        _validate_record_envelope(value, case_id=item["caseId"], layer=item["layer"])
        record_values.append(value)
        destination = _record_destination(item["layer"], item["caseId"])
        evidence_sources[(item["layer"], item["caseId"])] = (source_path, identity)
        predicted_records.append(
            {"caseId": item["caseId"], "layer": item["layer"], **_identity_at(identity, destination)}
        )
    existing_allowed = tuple(
        sorted(
            item.name
            for item in safe_source_staging.iterdir()
            if item.name not in PACKAGE_NAMES and item.is_dir() and not item.is_symlink()
        )
    )
    package_payload = candidate_evidence.capture_package_manifest(
        safe_source_staging, PACKAGE_NAMES, allowed_root_entries=existing_allowed
    )
    package_identity = _json_identity(package_payload, relative_path=contract.PACKAGE_MANIFEST_PATH)
    oracle_relative = str(cast(Mapping[str, object], source["oracleManifest"])["relativePath"])
    source_pointer = _with_role("source_manifest", governance[oracle_relative])
    wire_pointer = _with_role("wire_schema", governance[contract.WIRE_SCHEMA_PATH])
    package_pointer = _with_role("package_manifest", package_identity)
    manifest_sources: dict[str, tuple[Path, dict[str, object]]] = {}
    for layer, raw_path in planned_manifests.items():
        if raw_path is None:
            continue
        source_path = safe_evidence / raw_path
        value, identity = contract.read_json_object_with_identity(
            source_path,
            relative_to=safe_evidence,
            label=f"{layer}_manifest",
        )
        layer_records = [item for item in predicted_records if item["layer"] == layer]
        expected = _manifest_expected(
            layer=layer,
            records=layer_records,
            source_pointer=source_pointer,
            wire_pointer=wire_pointer,
            package_pointer=package_pointer,
        )
        if value != expected:
            raise V4StagingBuildError(f"evidence_manifest_not_closed_or_mismatched:{layer}")
        manifest_sources[layer] = (source_path, identity)
    pointer_by_layer: dict[str, dict[str, object]] = {
        "source_oracle": source_pointer,
        "machine_wire": wire_pointer,
        "packaged": package_pointer,
    }
    for layer, (_, identity) in manifest_sources.items():
        role = contract.POINTER_ROLE_BY_LAYER[layer]
        pointer_by_layer[layer] = _with_role(role, _identity_at(identity, _manifest_destination(layer)))
    index_records = [
        {**item, "identityPointers": [pointer_by_layer[cast(str, item["layer"])]]} for item in predicted_records
    ]
    index = {
        "schema": contract.EVIDENCE_SCHEMA,
        "candidateId": candidate_id,
        "layerStatus": statuses,
        "records": index_records,
    }
    contract.validate_evidence_index(index, candidate_id=candidate_id)
    contract.validate_json_schema(index, docwen_repo / contract.EVIDENCE_SCHEMA_PATH)
    try:
        fixtures = evidence_contract.source_fixture_identities(record_values)
        artifact_identities = evidence_contract.evidence_artifact_identities(record_values)
    except evidence_contract.V4EvidenceContractError as exc:
        raise V4StagingBuildError(str(exc)) from exc
    for fixture in fixtures:
        relative = _safe_relative(fixture.get("relativePath"), label="source_fixture")
        if contract.file_identity(docwen_repo / relative, relative_to=docwen_repo) != fixture:
            raise V4StagingBuildError(f"source_fixture_bytes_drifted:{relative}")
    evidence_artifacts: list[tuple[str, dict[str, object]]] = []
    for identity in artifact_identities:
        destination = cast(str, identity["relativePath"])
        source_relative = destination.removeprefix("evidence/")
        artifact_source_identity = _identity_at(identity, source_relative)
        if (
            contract.file_identity(safe_evidence / source_relative, relative_to=safe_evidence)
            != artifact_source_identity
        ):
            raise V4StagingBuildError(f"evidence_artifact_bytes_drifted:{destination}")
        evidence_artifacts.append((source_relative, artifact_source_identity))
    source_recheck, source_recheck_identity = _read_external(
        source_checkpoint, expected_sha256=source_checkpoint_sha256, label="source_checkpoint"
    )
    plan_recheck, plan_recheck_identity = _read_external(
        evidence_plan, expected_sha256=evidence_plan_sha256, label="evidence_plan"
    )
    if source_recheck != source or source_recheck_identity != source_identity:
        raise V4StagingBuildError("source_checkpoint_changed_before_staging")
    if plan_recheck != plan or plan_recheck_identity != plan_identity:
        raise V4StagingBuildError("evidence_plan_changed_before_staging")
    contract.verify_source_checkpoint(source_recheck, docwen_repo=docwen_repo)
    checkpoint_publish_started = False
    checkpoint_result: dict[str, object]
    try:
        sealer._copy_stable(safe_source_staging, preparing)
        if contract.capture_tree_stable(safe_source_staging) != source_tree:
            raise V4StagingBuildError("source_staging_changed_during_build")
        for relative, identity in governance.items():
            _copy_checked(
                docwen_repo / relative,
                preparing / relative,
                source_root=docwen_repo,
                destination_root=preparing,
                expected=identity,
            )
        for fixture in fixtures:
            relative = cast(str, fixture["relativePath"])
            _copy_checked(
                docwen_repo / relative,
                preparing / relative,
                source_root=docwen_repo,
                destination_root=preparing,
                expected=fixture,
            )
        for item in planned_records:
            source_path, identity = evidence_sources[(item["layer"], item["caseId"])]
            _copy_checked(
                source_path,
                preparing / _record_destination(item["layer"], item["caseId"]),
                source_root=safe_evidence,
                destination_root=preparing,
                expected=identity,
            )
        for source_relative, identity in evidence_artifacts:
            _copy_checked(
                safe_evidence / source_relative,
                preparing / f"evidence/{source_relative}",
                source_root=safe_evidence,
                destination_root=preparing,
                expected=identity,
            )
        for layer, (source_path, identity) in manifest_sources.items():
            _copy_checked(
                source_path,
                preparing / _manifest_destination(layer),
                source_root=safe_evidence,
                destination_root=preparing,
                expected=identity,
            )
        (preparing / "_evidence").mkdir()
        allowed = tuple(
            sorted(
                item.name
                for item in preparing.iterdir()
                if item.name not in PACKAGE_NAMES and item.is_dir() and not item.is_symlink()
            )
        )
        actual_package = candidate_evidence.write_package_manifest(
            preparing,
            PACKAGE_NAMES,
            preparing / contract.PACKAGE_MANIFEST_PATH,
            allowed_root_entries=allowed,
        )
        if actual_package != package_payload:
            raise V4StagingBuildError("package_manifest_changed_while_staging")
        contract.write_json_exclusive(preparing / "_evidence/v4-evidence-index.json", index)
        contract.validate_json_schema(index, preparing / contract.EVIDENCE_SCHEMA_PATH)
        stable_tree = contract.capture_tree_stable(preparing)
        contract.verify_index_files(preparing, index_records, stable_tree=stable_tree)
        source_final, source_final_identity = _read_external(
            source_checkpoint, expected_sha256=source_checkpoint_sha256, label="source_checkpoint"
        )
        plan_final, plan_final_identity = _read_external(
            evidence_plan, expected_sha256=evidence_plan_sha256, label="evidence_plan"
        )
        if source_final != source or source_final_identity != source_identity:
            raise V4StagingBuildError("source_checkpoint_changed_before_publish")
        if plan_final != plan or plan_final_identity != plan_identity:
            raise V4StagingBuildError("evidence_plan_changed_before_publish")
        contract.verify_source_checkpoint(source_final, docwen_repo=docwen_repo)
        _preflight_source_files(docwen_repo, governance)
        preparing.rename(final_root)
        if contract.capture_tree_stable(final_root) != stable_tree:
            raise V4StagingBuildError("staging_tree_changed_after_publish")
        raw_checkpoint_result = sealer.write_staging_checkpoint(
            docwen_repo=docwen_repo,
            staging_root=final_root,
            source_checkpoint=source_checkpoint,
            source_checkpoint_sha256=source_checkpoint_sha256,
            output=checkpoint_preparing,
        )
        checkpoint_publish_started = True
        checkpoint_preparing.rename(safe_checkpoint)
        checkpoint_identity = cast(
            dict[str, object], contract.file_identity(safe_checkpoint, relative_to=safe_checkpoint.parent)
        )
        checkpoint_result = {"checkpoint": checkpoint_identity}
        raw_identity = cast(Mapping[str, object], raw_checkpoint_result["checkpoint"])
        if any(checkpoint_identity[key] != raw_identity[key] for key in ("bytes", "sha256")):
            raise V4StagingBuildError("staging_checkpoint_changed_during_publish")
        checkpoint_payload = contract.read_json_object(safe_checkpoint, label="published_staging_checkpoint")
        if checkpoint_payload.get("stagingTree") != sealer._payload_tree(final_root):
            raise V4StagingBuildError("staging_changed_after_checkpoint_publish")
        if (
            contract.file_identity(safe_checkpoint, relative_to=safe_checkpoint.parent)
            != checkpoint_result["checkpoint"]
        ):
            raise V4StagingBuildError("staging_checkpoint_changed_after_publish")
        contract.verify_source_checkpoint(source, docwen_repo=docwen_repo)
    except BaseException as proof_error:
        quarantine_errors: list[BaseException] = []
        try:
            _quarantine(
                preparing if preparing.exists() else final_root,
                output_root=safe_output,
                candidate_id=candidate_id,
            )
        except BaseException as exc:
            quarantine_errors.append(exc)
        owned_checkpoint = checkpoint_preparing
        if checkpoint_publish_started and not checkpoint_preparing.exists() and safe_checkpoint.exists():
            owned_checkpoint = safe_checkpoint
        if owned_checkpoint.exists():
            rejected_checkpoint = safe_checkpoint.parent / f".rejected-{safe_checkpoint.name}-{uuid.uuid4().hex}"
            try:
                owned_checkpoint.rename(rejected_checkpoint)
            except BaseException as exc:
                quarantine_errors.append(exc)
        if quarantine_errors:
            error = V4StagingBuildError("staging_failure_and_quarantine_failed:canonical_absent_or_nonreusable")
            error.add_note(f"staging failure: {proof_error}")
            for quarantine_error in quarantine_errors:
                error.add_note(f"quarantine failure: {quarantine_error}")
            raise error from quarantine_errors[0]
        raise
    return {
        "candidateId": candidate_id,
        "stagingRoot": str(final_root),
        "stagingTreeManifestSha256": contract.capture_tree_stable(final_root)["manifestSha256"],
        "sourceCheckpointSha256": source_checkpoint_sha256,
        "evidencePlan": plan_identity,
        "stagingCheckpoint": checkpoint_result["checkpoint"],
    }


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Assemble already-produced v4 evidence into an immutable staging tree."
    )
    parser.add_argument("--docwen-repo", type=Path, required=True)
    parser.add_argument("--source-staging", type=Path, required=True)
    parser.add_argument("--evidence-root", type=Path, required=True)
    parser.add_argument("--evidence-plan", type=Path, required=True)
    parser.add_argument("--evidence-plan-sha256", required=True)
    parser.add_argument("--source-checkpoint", type=Path, required=True)
    parser.add_argument("--source-checkpoint-sha256", required=True)
    parser.add_argument("--staging-checkpoint-output", type=Path, required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    parser.add_argument("--candidate-id", required=True)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        result = build_staging(
            docwen_repo=args.docwen_repo,
            source_staging=args.source_staging,
            evidence_root=args.evidence_root,
            evidence_plan=args.evidence_plan,
            evidence_plan_sha256=args.evidence_plan_sha256,
            source_checkpoint=args.source_checkpoint,
            source_checkpoint_sha256=args.source_checkpoint_sha256,
            staging_checkpoint_output=args.staging_checkpoint_output,
            output_root=args.output_root,
            candidate_id=args.candidate_id,
        )
    except (
        V4StagingBuildError,
        contract.V4CandidateContractError,
        sealer.V4CandidateSealError,
        candidate_evidence.EvidenceError,
        OSError,
    ) as exc:
        print(f"build_v4_candidate_staging_error:{exc}", file=sys.stderr)
        return 2
    print(json.dumps({"success": True, "data": result}, ensure_ascii=False, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
