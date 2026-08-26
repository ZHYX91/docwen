from __future__ import annotations

import re
from collections import Counter
from collections.abc import Mapping
from pathlib import Path
from typing import Any, cast

from scripts.release import candidate_evidence
from scripts.release import v4_candidate_contract as contract
from scripts.release import v4_evidence_contract as evidence_contract
from scripts.release import v4_evidence_io as evidence_io
from scripts.release import v4_package_input_contract as input_contract


def _write_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    contract.write_json_exclusive(path, value)


def _identity(path: Path, *, root: Path) -> dict[str, object]:
    try:
        return evidence_io.file_identity(path, relative_to=root)
    except evidence_io.V4EvidenceContractError as exc:
        raise input_contract.V4PackageInputError(str(exc)) from exc


def _read_json(path: Path, *, root: Path, label: str) -> tuple[dict[str, Any], dict[str, object]]:
    try:
        return evidence_io.read_json_object(path, relative_to=root, label=label)
    except evidence_io.V4EvidenceContractError as exc:
        raise input_contract.V4PackageInputError(str(exc)) from exc


def _safe_relative(value: object, *, label: str) -> str:
    try:
        return evidence_contract.validate_relative_path(value, label=label)
    except evidence_contract.V4EvidenceContractError as exc:
        raise input_contract.V4PackageInputError(str(exc)) from exc


def _oracle_case(repo: Path, case_id: str) -> tuple[dict[str, Any], dict[str, object], bytes]:
    relative = input_contract.SOURCE_ORACLE_ROOT / "corpus" / f"{case_id}.case.json"
    fixture, identity = _read_json(repo / relative, root=repo, label=f"source_oracle:{case_id}")
    if fixture.get("case_id") != case_id or not isinstance(fixture.get("source"), str):
        raise input_contract.V4PackageInputError(f"source_oracle_case_invalid:{case_id}")
    return fixture, identity, cast(str, fixture["source"]).encode("utf-8")


def _source_payload(repo: Path, case_id: str, dimension: str | None) -> tuple[dict[str, object], bytes]:
    fixture, identity, source = _oracle_case(repo, case_id)
    payload = {
        "schema": "docwen.v4_source_oracle_observation.v1",
        "fixture": identity,
        "sourceSha256": input_contract.sha256_bytes(source),
        "expectedDiagnostics": evidence_contract._normalized_fixture_diagnostics(fixture),
        "invalidIdDimension": dimension,
    }
    if not evidence_contract.payload_shape_matches("source_oracle", payload):
        raise input_contract.V4PackageInputError(f"source_oracle_payload_invalid:{case_id}")
    return payload, source


def _evidence_artifact(root: Path, relative: str, payload: bytes | object) -> dict[str, object]:
    safe_relative = _safe_relative(relative, label="evidence_artifact")
    path = root / safe_relative
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists() or path.is_symlink():
        raise input_contract.V4PackageInputError(f"evidence_artifact_exists:{safe_relative}")
    path.write_bytes(payload if isinstance(payload, bytes) else input_contract.json_bytes(payload))
    identity = _identity(path, root=root)
    return {
        "relativePath": f"evidence/{identity['relativePath']}",
        "bytes": identity["bytes"],
        "sha256": identity["sha256"],
    }


def _record_ref(record: Mapping[str, object]) -> dict[str, object]:
    return {
        "caseId": record["caseId"],
        "layer": record["layer"],
        **evidence_contract.identity_core(record),
    }


def _source_pointer(checkpoint: Mapping[str, object]) -> dict[str, object]:
    oracle = checkpoint.get("oracleManifest")
    if not isinstance(oracle, dict):
        raise input_contract.V4PackageInputError("checkpoint_oracle_manifest_missing")
    return evidence_contract.pointer_identity("source_manifest", oracle)


def _source_manifest(checkpoint: Mapping[str, object], repo: Path) -> tuple[Path, list[dict[str, object]]]:
    identity = checkpoint.get("oracleManifest")
    if not isinstance(identity, dict):
        raise input_contract.V4PackageInputError("checkpoint_oracle_manifest_missing")
    relative = _safe_relative(identity.get("relativePath"), label="source_manifest")
    manifest, actual = _read_json(repo / relative, root=repo, label="source_manifest")
    raw_files = manifest.get("files")
    if actual != identity or not isinstance(raw_files, list):
        raise input_contract.V4PackageInputError("source_manifest_identity_or_files_invalid")
    if not all(isinstance(item, dict) for item in raw_files):
        raise input_contract.V4PackageInputError("source_manifest_files_invalid")
    return Path(relative).parent, cast(list[dict[str, object]], raw_files)


def _require_manifest_bound(
    fixture: Mapping[str, object], *, manifest_base: Path, manifest_files: list[dict[str, object]], case_id: str
) -> None:
    try:
        relative = Path(str(fixture["relativePath"])).relative_to(manifest_base).as_posix()
    except (KeyError, ValueError) as exc:
        raise input_contract.V4PackageInputError(f"source_fixture_outside_oracle:{case_id}") from exc
    matches = [item for item in manifest_files if item.get("path") == relative]
    if len(matches) != 1 or any(matches[0].get(key) != fixture.get(key) for key in ("bytes", "sha256")):
        raise input_contract.V4PackageInputError(f"source_fixture_not_manifest_bound:{case_id}")


def _wire_pointer(checkpoint: Mapping[str, object]) -> dict[str, object]:
    final_spec = checkpoint.get("finalSpecIdentity")
    files = final_spec.get("files") if isinstance(final_spec, dict) else None
    matches = (
        [item for item in files if isinstance(item, dict) and item.get("relativePath") == contract.WIRE_SCHEMA_PATH]
        if isinstance(files, list)
        else []
    )
    if len(matches) != 1:
        raise input_contract.V4PackageInputError("checkpoint_wire_schema_identity_missing")
    return evidence_contract.pointer_identity("wire_schema", matches[0])


def build_evidence(
    *,
    evidence_root: Path,
    package_root: Path,
    docwen_clone: Path,
    checkpoint: Mapping[str, object],
    candidate_id: str,
    harness: input_contract.HarnessInput,
    harness_output: input_contract.HarnessOutput,
    executable_identity: Mapping[str, object],
    package_names: tuple[str, str],
    plan_schema: str,
) -> tuple[dict[str, object], dict[str, object], dict[str, object]]:
    transcript = input_contract.validate_session_transcript(
        harness_output.transcript,
        harness=harness,
        outputs=harness_output.cases,
    )
    streams = cast(Mapping[str, object], transcript["streams"])
    if streams.get("requestSha256") != harness_output.request_digest:
        raise input_contract.V4PackageInputError("machine_transcript_request_digest_mismatch")
    records: list[dict[str, str]] = []
    predicted: list[dict[str, object]] = []

    def add(case_id: str, layer: str, payload: Mapping[str, object]) -> dict[str, object]:
        if not re.fullmatch(r"[a-z0-9][a-z0-9.-]{0,127}", case_id):
            raise input_contract.V4PackageInputError(f"evidence_case_id_invalid:{case_id}")
        envelope = {
            "schema": evidence_contract.RECORD_SCHEMA,
            "caseId": case_id,
            "layer": layer,
            "result": "passed",
            "observation": {"kind": layer, "payload": dict(payload)},
        }
        if not evidence_contract.record_envelope_matches(envelope, case_id=case_id, layer=layer):
            raise input_contract.V4PackageInputError(f"evidence_record_invalid:{layer}:{case_id}")
        relative = f"records/{layer}/{case_id}.json"
        path = evidence_root / relative
        _write_json(path, envelope)
        source_identity = _identity(path, root=evidence_root)
        record = {
            "caseId": case_id,
            "layer": layer,
            "relativePath": evidence_contract.record_path(layer, case_id),
            "bytes": source_identity["bytes"],
            "sha256": source_identity["sha256"],
        }
        records.append({"caseId": case_id, "layer": layer, "artifact": relative})
        predicted.append(record)
        return record

    source_records: dict[str, dict[str, object]] = {}
    source_bytes: dict[str, bytes] = {}
    source_payloads: dict[str, dict[str, object]] = {}
    manifest_base, manifest_files = _source_manifest(checkpoint, docwen_clone)

    def add_source(case_id: str, dimension: str | None) -> None:
        payload, raw = _source_payload(docwen_clone, case_id, dimension)
        _require_manifest_bound(
            cast(Mapping[str, object], payload["fixture"]),
            manifest_base=manifest_base,
            manifest_files=manifest_files,
            case_id=case_id,
        )
        source_records[case_id] = add(case_id, "source_oracle", payload)
        source_bytes[case_id] = raw
        source_payloads[case_id] = payload

    for dimension, case_id in contract.REQUIRED_INVALID_CASES.items():
        add_source(case_id, dimension)
    for case_id in sorted({case.source_oracle_case_id for case in harness.cases}):
        if case_id in source_records:
            continue
        add_source(case_id, None)
    for case in harness.cases:
        if source_bytes[case.source_oracle_case_id] != case.source_bytes:
            raise input_contract.V4PackageInputError(f"harness_source_bytes_changed:{case.case_id}")

    terminal = harness_output.validation_terminal
    if not evidence_contract.payload_shape_matches(
        "machine_wire",
        {
            "schema": "docwen.v4_machine_wire_observation.v1",
            "protocol": "docwen.machine.v1",
            "transcript": {"relativePath": "evidence/artifacts/x", "bytes": 1, "sha256": "0" * 64},
            "terminal": terminal,
            "terminalSha256": evidence_contract._payload_hash(terminal),
        },
    ):
        raise input_contract.V4PackageInputError("validation_terminal_not_closed_wire_evidence")
    wire_terminal_artifact = _evidence_artifact(
        evidence_root,
        f"artifacts/{input_contract.HARNESS_CASE_PREFIX}/validation-terminal.json",
        terminal,
    )
    wire_payload = {
        "schema": "docwen.v4_machine_wire_observation.v1",
        "protocol": "docwen.machine.v1",
        "transcript": wire_terminal_artifact,
        "terminal": terminal,
        "terminalSha256": evidence_contract._payload_hash(terminal),
    }
    wire_record = add(f"{input_contract.HARNESS_CASE_PREFIX}-wire", "machine_wire", wire_payload)
    normalized_wire = [
        {
            "severity": item["severity"],
            "code": item["code"],
            "range": item["range"],
            "relatedRanges": item["related_ranges"],
        }
        for item in cast(list[dict[str, object]], cast(Mapping[str, object], terminal["params"])["diagnostics"])
    ]
    if normalized_wire != source_payloads[input_contract.VALIDATION_CASE_ID]["expectedDiagnostics"]:
        raise input_contract.V4PackageInputError("source_wire_diagnostics_not_equal")
    comparison_record = add(
        f"{input_contract.HARNESS_CASE_PREFIX}-wire-eq",
        "source_wire_comparison",
        {
            "schema": "docwen.v4_source_wire_comparison_observation.v1",
            "result": "equal",
            "sourceRecord": _record_ref(source_records[input_contract.VALIDATION_CASE_ID]),
            "wireRecord": _record_ref(wire_record),
            "comparedFields": ["diagnostics"],
            "mismatches": [],
        },
    )
    package_manifest = candidate_evidence.capture_package_manifest(
        package_root,
        package_names,
        allowed_root_entries=(),
    )
    package_manifest_raw = contract.json_bytes(package_manifest)
    package_identity = {
        "relativePath": contract.PACKAGE_MANIFEST_PATH,
        "bytes": len(package_manifest_raw),
        "sha256": input_contract.sha256_bytes(package_manifest_raw),
    }
    package_files = package_manifest.get("files")
    executable_relative = Path(str(executable_identity.get("relativePath", "")))
    executable_entry = {
        "package": executable_relative.parts[0] if executable_relative.parts else "",
        "path": Path(*executable_relative.parts[1:]).as_posix(),
        "bytes": executable_identity.get("bytes"),
        "sha256": executable_identity.get("sha256"),
    }
    if (
        not isinstance(package_files, list)
        or sum(
            isinstance(item, dict) and all(item.get(key) == value for key, value in executable_entry.items())
            for item in package_files
        )
        != 1
    ):
        raise input_contract.V4PackageInputError("packaged_executable_not_manifest_bound")
    machine_stdout = _evidence_artifact(
        evidence_root,
        f"artifacts/{input_contract.HARNESS_CASE_PREFIX}/machine-stdio-transcript.json",
        harness_output.transcript,
    )
    machine_stderr = _evidence_artifact(
        evidence_root,
        f"artifacts/{input_contract.HARNESS_CASE_PREFIX}/machine-stderr.bin",
        harness_output.stderr,
    )
    package_record = add(
        f"{input_contract.HARNESS_CASE_PREFIX}-pkg",
        "packaged",
        {
            "schema": "docwen.v4_packaged_observation.v1",
            "packageManifest": evidence_contract.identity_core(package_identity),
            "executable": dict(executable_identity),
            "invocation": {
                "argv": [executable_identity["relativePath"], "serve", "--stdio"],
                "exitCode": 0,
                "stdout": machine_stdout,
                "stderr": machine_stderr,
            },
        },
    )
    output_by_id = {case.case_id: case for case in harness_output.cases}
    expected_ids = [case.case_id for case in harness.cases]
    if list(output_by_id) != expected_ids or len(output_by_id) != len(harness_output.cases):
        raise input_contract.V4PackageInputError("harness_output_case_set_or_order_invalid")
    case_summaries: list[dict[str, object]] = []
    for case in harness.cases:
        output = output_by_id[case.case_id]
        if output.roundtrip != case.expected_roundtrip:
            raise input_contract.V4PackageInputError(f"exact_two_roundtrip_not_byte_exact:{case.case_id}")
        verified_inspection = input_contract.inspect_docx(
            output.docx,
            case.expected_ooxml,
            case.neutral_envelope,
            case.plan_envelope,
        )
        if output.inspection != verified_inspection:
            raise input_contract.V4PackageInputError(f"harness_output_inspection_mismatch:{case.case_id}")
        artifact_root = f"artifacts/{input_contract.HARNESS_CASE_PREFIX}/{case.case_id}"
        roundtrip_input = _evidence_artifact(evidence_root, f"{artifact_root}/roundtrip-input.md", case.source_bytes)
        roundtrip_output = _evidence_artifact(evidence_root, f"{artifact_root}/roundtrip-output.md", output.roundtrip)
        roundtrip_record = add(
            f"{input_contract.HARNESS_CASE_PREFIX}-{case.case_id}-rt",
            "roundtrip",
            {
                "schema": "docwen.v4_roundtrip_observation.v1",
                "sourceRecord": _record_ref(source_records[case.source_oracle_case_id]),
                "packageRecord": _record_ref(package_record),
                "input": roundtrip_input,
                "output": roundtrip_output,
                "byteExact": True,
            },
        )
        docx_artifact = _evidence_artifact(evidence_root, f"{artifact_root}/headless.docx", output.docx)
        headless_record = add(
            f"{input_contract.HARNESS_CASE_PREFIX}-{case.case_id}-xml",
            "headless_ooxml",
            {
                "schema": "docwen.v4_headless_ooxml_observation.v1",
                "packageRecord": _record_ref(package_record),
                "artifact": docx_artifact,
                "inspection": output.inspection,
            },
        )
        case_summaries.append(
            {
                "caseId": case.case_id,
                "sourceOracleCaseId": case.source_oracle_case_id,
                "roundtrip": roundtrip_record,
                "headless": headless_record,
            }
        )

    source_pointer = _source_pointer(checkpoint)
    wire_pointer = _wire_pointer(checkpoint)
    package_pointer = evidence_contract.pointer_identity("package_manifest", package_identity)
    manifests: dict[str, str | None] = {}
    for layer in evidence_contract.EXTERNAL_MANIFEST_LAYERS:
        if layer in evidence_contract.HOST_LAYERS:
            manifests[layer] = None
            continue
        layer_records = [item for item in predicted if item["layer"] == layer]
        relative = f"manifests/{layer}.json"
        value = evidence_contract.manifest_expected(
            layer=layer,
            records=layer_records,
            source_pointer=source_pointer,
            wire_pointer=wire_pointer,
            package_pointer=package_pointer,
        )
        _write_json(evidence_root / relative, value)
        manifests[layer] = relative
    statuses = {
        **dict.fromkeys(contract.REQUIRED_LAYERS[:6], "passed"),
        **dict.fromkeys(contract.REQUIRED_LAYERS[6:], "not_run"),
    }
    plan = {
        "schema": plan_schema,
        "candidateId": candidate_id,
        "layerStatus": statuses,
        "records": records,
        "manifests": manifests,
    }
    if not all(
        item["layer"] == "source_oracle" or item["caseId"].startswith(input_contract.HARNESS_CASE_PREFIX)
        for item in records
    ):
        raise input_contract.V4PackageInputError("evidence_plan_harness_identity_missing")
    input_contract.reject_legacy_harness(
        {"harness": {"id": input_contract.HARNESS_ID, "version": input_contract.HARNESS_VERSION}, "records": records}
    )
    expected_counts = Counter(
        {
            "source_oracle": len(source_records),
            "machine_wire": 1,
            "source_wire_comparison": 1,
            "packaged": 1,
            "roundtrip": len(harness.cases),
            "headless_ooxml": len(harness.cases),
        }
    )
    if Counter(item["layer"] for item in records) != expected_counts or len(predicted) != len(records):
        raise input_contract.V4PackageInputError("evidence_record_counts_not_exact")
    return (
        plan,
        package_manifest,
        {
            "comparison": comparison_record,
            "package": package_record,
            "cases": case_summaries,
        },
    )
