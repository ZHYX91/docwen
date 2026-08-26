from __future__ import annotations

import hashlib
import json
from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import cast

from scripts.release import v4_evidence_io as evidence_io

V4EvidenceContractError = evidence_io.V4EvidenceContractError
validate_relative_path = evidence_io.validate_relative_path
safe_regular_file = evidence_io.safe_regular_file
file_identity = evidence_io.file_identity
sha256_file = evidence_io.sha256_file
read_json_object = evidence_io.read_json_object
capture_tree_stable = evidence_io.capture_tree_stable

RECORD_SCHEMA = "docwen.v4_evidence_record.v1"
POINTER_ROLE_BY_LAYER = {
    "source_oracle": "source_manifest",
    "machine_wire": "wire_schema",
    "source_wire_comparison": "comparison_manifest",
    "packaged": "package_manifest",
    "roundtrip": "roundtrip_manifest",
    "headless_ooxml": "headless_ooxml_manifest",
    "word_host": "host_manifest",
    "wps_host": "host_manifest",
    "libreoffice_host": "host_manifest",
}
EXTERNAL_MANIFEST_LAYERS = (
    "source_wire_comparison",
    "roundtrip",
    "headless_ooxml",
    "word_host",
    "wps_host",
    "libreoffice_host",
)
HOST_LAYERS = ("word_host", "wps_host", "libreoffice_host")
_HEX64 = frozenset("0123456789abcdef")
_INVALID_DIMENSIONS = {"dot", "slash", "underscore", "over128", "duplicate", "kindMismatch"}


def record_path(layer: str, case_id: str) -> str:
    return f"evidence/records/{layer}/{case_id}.json"


def manifest_path(layer: str) -> str:
    return f"evidence/manifests/{layer}.json"


def identity_core(value: Mapping[str, object]) -> dict[str, object]:
    return {key: value[key] for key in ("relativePath", "bytes", "sha256")}


def pointer_identity(role: str, identity: Mapping[str, object]) -> dict[str, object]:
    return {"role": role, **identity_core(identity)}


def _is_hex64(value: object) -> bool:
    return isinstance(value, str) and len(value) == 64 and set(value) <= _HEX64


def _identity_shape(value: object) -> bool:
    return bool(
        isinstance(value, dict)
        and set(value) == {"relativePath", "bytes", "sha256"}
        and isinstance(value.get("relativePath"), str)
        and isinstance(value.get("bytes"), int)
        and not isinstance(value.get("bytes"), bool)
        and cast(int, value["bytes"]) >= 0
        and _is_hex64(value.get("sha256"))
    )


def _record_ref_shape(value: object, *, layer: str | None = None) -> bool:
    return bool(
        isinstance(value, dict)
        and set(value) == {"caseId", "layer", "relativePath", "bytes", "sha256"}
        and isinstance(value.get("caseId"), str)
        and isinstance(value.get("layer"), str)
        and (layer is None or value.get("layer") == layer)
        and _identity_shape({key: value[key] for key in ("relativePath", "bytes", "sha256")})
    )


def _range_shape(value: object) -> bool:
    return bool(
        isinstance(value, dict)
        and set(value) == {"start", "end"}
        and all(isinstance(value.get(key), int) and not isinstance(value.get(key), bool) for key in ("start", "end"))
        and 0 <= cast(int, value["start"]) < cast(int, value["end"])
    )


def _diagnostic_shape(value: object) -> bool:
    return bool(
        isinstance(value, dict)
        and set(value) == {"severity", "code", "range", "relatedRanges"}
        and value.get("severity") in {"error", "warning"}
        and isinstance(value.get("code"), str)
        and bool(value["code"])
        and _range_shape(value.get("range"))
        and isinstance(value.get("relatedRanges"), list)
        and all(_range_shape(item) for item in cast(list[object], value["relatedRanges"]))
    )


def _wire_source_shape(value: object) -> bool:
    return bool(
        isinstance(value, dict)
        and set(value) == {"input_id", "sha256", "encoding", "coordinate_system", "offset_base", "range_end"}
        and isinstance(value.get("input_id"), str)
        and bool(value["input_id"])
        and _is_hex64(value.get("sha256"))
        and value.get("encoding") == "utf-8"
        and value.get("coordinate_system") == "unicode_code_point"
        and value.get("offset_base") == 0
        and value.get("range_end") == "exclusive"
    )


def _edit_range_shape(value: object) -> bool:
    return bool(
        isinstance(value, dict)
        and set(value) == {"start", "end"}
        and all(isinstance(value.get(key), int) and not isinstance(value.get(key), bool) for key in ("start", "end"))
        and 0 <= cast(int, value["start"]) <= cast(int, value["end"])
    )


def _wire_fix_shape(value: object) -> bool:
    if not isinstance(value, dict) or set(value) != {"fix_id", "edits"}:
        return False
    edits = value.get("edits")
    return bool(
        isinstance(value.get("fix_id"), str)
        and value["fix_id"]
        and isinstance(edits, list)
        and edits
        and all(
            isinstance(edit, dict)
            and set(edit) == {"range", "replacement"}
            and _edit_range_shape(edit.get("range"))
            and isinstance(edit.get("replacement"), str)
            for edit in edits
        )
    )


def _wire_diagnostic_shape(value: object) -> bool:
    if not isinstance(value, dict):
        return False
    related = value.get("related_ranges")
    fixes = value.get("fixes")
    return bool(
        set(value) == {"severity", "code", "message", "evidence_schema", "source", "range", "related_ranges", "fixes"}
        and value.get("severity") in {"error", "warning"}
        and isinstance(value.get("code"), str)
        and bool(value["code"])
        and isinstance(value.get("message"), str)
        and bool(value["message"])
        and value.get("evidence_schema") == "docwen.machine.diagnostic_evidence.v1"
        and _wire_source_shape(value.get("source"))
        and _range_shape(value.get("range"))
        and isinstance(related, list)
        and all(_range_shape(item) for item in related)
        and isinstance(fixes, list)
        and all(_wire_fix_shape(item) for item in fixes)
    )


def _bundle_shape(value: object) -> bool:
    if not isinstance(value, dict) or set(value) != {
        "schema",
        "layout_schema",
        "bundle_id",
        "task_id",
        "producer",
        "artifacts",
        "entries",
        "relations",
    }:
        return False
    producer = value.get("producer")
    artifacts = value.get("artifacts")
    entries = value.get("entries")
    return bool(
        value.get("schema") == "docwen.artifact_bundle.v2"
        and value.get("layout_schema") in {"docwen.artifact_layout.v1", "docwen.document_node.v1"}
        and isinstance(value.get("bundle_id"), str)
        and bool(value["bundle_id"])
        and isinstance(value.get("task_id"), str)
        and bool(value["task_id"])
        and isinstance(producer, dict)
        and set(producer) == {"name", "product_version", "machine_protocol"}
        and producer.get("name") == "DocWen"
        and producer.get("product_version") == "0.9.0"
        and producer.get("machine_protocol") == "docwen.machine.v1"
        and isinstance(artifacts, list)
        and bool(artifacts)
        and all(
            isinstance(item, dict)
            and set(item)
            == {
                "artifact_id",
                "kind",
                "locator",
                "logical_path",
                "suggested_name",
                "media_type",
                "size_bytes",
                "sha256",
            }
            and all(
                isinstance(item.get(key), str) and item[key]
                for key in ("artifact_id", "kind", "locator", "logical_path", "suggested_name", "media_type")
            )
            and isinstance(item.get("size_bytes"), int)
            and not isinstance(item.get("size_bytes"), bool)
            and cast(int, item["size_bytes"]) >= 0
            and _is_hex64(item.get("sha256"))
            for item in artifacts
        )
        and isinstance(entries, list)
        and bool(entries)
        and all(
            isinstance(item, dict)
            and set(item) == {"artifact_id", "role", "ordinal", "preferred"}
            and isinstance(item.get("artifact_id"), str)
            and bool(item["artifact_id"])
            and item.get("role") in {"primary", "supplementary"}
            and isinstance(item.get("ordinal"), int)
            and not isinstance(item.get("ordinal"), bool)
            and cast(int, item["ordinal"]) >= 0
            and isinstance(item.get("preferred"), bool)
            for item in entries
        )
        and value.get("relations") == []
    )


def _terminal_shape(value: object) -> bool:
    if not isinstance(value, dict) or set(value) != {"jsonrpc", "method", "params"}:
        return False
    params = value.get("params")
    if not isinstance(params, dict) or set(params) != {"task_id", "sequence", "bundle", "diagnostics", "metrics"}:
        return False
    diagnostics = params.get("diagnostics")
    metrics = params.get("metrics")
    bundle = params.get("bundle")
    return bool(
        value.get("jsonrpc") == "2.0"
        and value.get("method") == "task/completed"
        and isinstance(params.get("task_id"), str)
        and bool(params["task_id"])
        and isinstance(params.get("sequence"), int)
        and not isinstance(params.get("sequence"), bool)
        and cast(int, params["sequence"]) >= 1
        and _bundle_shape(bundle)
        and cast(dict[str, object], bundle)["task_id"] == params["task_id"]
        and isinstance(diagnostics, list)
        and bool(diagnostics)
        and all(_wire_diagnostic_shape(item) for item in diagnostics)
        and isinstance(metrics, dict)
        and set(metrics) == {"duration_ms", "input_bytes", "output_bytes"}
        and all(
            isinstance(metrics.get(key), int)
            and not isinstance(metrics.get(key), bool)
            and cast(int, metrics[key]) >= 0
            for key in metrics
        )
    )


def _payload_hash(value: object) -> str:
    raw = json.dumps(value, ensure_ascii=False, separators=(",", ":"), sort_keys=True).encode("utf-8")
    return hashlib.sha256(raw).hexdigest()


def payload_shape_matches(layer: str, payload: Mapping[str, object]) -> bool:
    if layer == "source_oracle":
        diagnostics = payload.get("expectedDiagnostics")
        dimension = payload.get("invalidIdDimension")
        return bool(
            set(payload) == {"schema", "fixture", "sourceSha256", "expectedDiagnostics", "invalidIdDimension"}
            and payload.get("schema") == "docwen.v4_source_oracle_observation.v1"
            and _identity_shape(payload.get("fixture"))
            and _is_hex64(payload.get("sourceSha256"))
            and isinstance(diagnostics, list)
            and all(_diagnostic_shape(item) for item in diagnostics)
            and (dimension is None or dimension in _INVALID_DIMENSIONS)
        )
    if layer == "machine_wire":
        terminal = payload.get("terminal")
        return bool(
            set(payload) == {"schema", "protocol", "transcript", "terminal", "terminalSha256"}
            and payload.get("schema") == "docwen.v4_machine_wire_observation.v1"
            and payload.get("protocol") == "docwen.machine.v1"
            and _identity_shape(payload.get("transcript"))
            and _terminal_shape(terminal)
            and payload.get("terminalSha256") == _payload_hash(terminal)
        )
    if layer == "source_wire_comparison":
        return bool(
            set(payload) == {"schema", "result", "sourceRecord", "wireRecord", "comparedFields", "mismatches"}
            and payload.get("schema") == "docwen.v4_source_wire_comparison_observation.v1"
            and payload.get("result") == "equal"
            and _record_ref_shape(payload.get("sourceRecord"), layer="source_oracle")
            and _record_ref_shape(payload.get("wireRecord"), layer="machine_wire")
            and payload.get("comparedFields") == ["diagnostics"]
            and payload.get("mismatches") == []
        )
    if layer == "packaged":
        invocation = payload.get("invocation")
        return bool(
            set(payload) == {"schema", "packageManifest", "executable", "invocation"}
            and payload.get("schema") == "docwen.v4_packaged_observation.v1"
            and _identity_shape(payload.get("packageManifest"))
            and _identity_shape(payload.get("executable"))
            and isinstance(invocation, dict)
            and set(invocation) == {"argv", "exitCode", "stdout", "stderr"}
            and isinstance(invocation.get("argv"), list)
            and bool(invocation["argv"])
            and all(isinstance(item, str) and item for item in cast(list[object], invocation["argv"]))
            and invocation.get("exitCode") == 0
            and _identity_shape(invocation.get("stdout"))
            and _identity_shape(invocation.get("stderr"))
        )
    if layer == "roundtrip":
        return bool(
            set(payload) == {"schema", "sourceRecord", "packageRecord", "input", "output", "byteExact"}
            and payload.get("schema") == "docwen.v4_roundtrip_observation.v1"
            and _record_ref_shape(payload.get("sourceRecord"), layer="source_oracle")
            and _record_ref_shape(payload.get("packageRecord"), layer="packaged")
            and _identity_shape(payload.get("input"))
            and _identity_shape(payload.get("output"))
            and payload.get("byteExact") is True
        )
    if layer == "headless_ooxml":
        inspection = payload.get("inspection")
        return bool(
            set(payload) == {"schema", "packageRecord", "artifact", "inspection"}
            and payload.get("schema") == "docwen.v4_headless_ooxml_observation.v1"
            and _record_ref_shape(payload.get("packageRecord"), layer="packaged")
            and _identity_shape(payload.get("artifact"))
            and isinstance(inspection, dict)
            and set(inspection) == {"bookmarkCount", "seqFieldCount", "refFieldCount", "violations"}
            and all(
                isinstance(inspection.get(key), int)
                and not isinstance(inspection.get(key), bool)
                and cast(int, inspection[key]) >= 0
                for key in ("bookmarkCount", "seqFieldCount", "refFieldCount")
            )
            and inspection.get("violations") == []
        )
    if layer in HOST_LAYERS:
        host = payload.get("host")
        return bool(
            set(payload) == {"schema", "packageRecord", "headlessRecord", "artifact", "host"}
            and payload.get("schema") == "docwen.v4_host_observation.v1"
            and _record_ref_shape(payload.get("packageRecord"), layer="packaged")
            and _record_ref_shape(payload.get("headlessRecord"), layer="headless_ooxml")
            and _identity_shape(payload.get("artifact"))
            and isinstance(host, dict)
            and set(host) == {"name", "version", "opened", "rendered", "saved", "violations"}
            and host.get("name") == layer.removesuffix("_host")
            and isinstance(host.get("version"), str)
            and bool(host["version"])
            and all(host.get(key) is True for key in ("opened", "rendered", "saved"))
            and host.get("violations") == []
        )
    return False


def record_envelope_matches(value: Mapping[str, object], *, case_id: object, layer: str) -> bool:
    observation = value.get("observation")
    return bool(
        set(value) == {"schema", "caseId", "layer", "result", "observation"}
        and value.get("schema") == RECORD_SCHEMA
        and value.get("caseId") == case_id
        and value.get("layer") == layer
        and value.get("result") == "passed"
        and isinstance(observation, dict)
        and set(observation) == {"kind", "payload"}
        and observation.get("kind") == layer
        and isinstance(observation.get("payload"), dict)
        and payload_shape_matches(layer, cast(Mapping[str, object], observation["payload"]))
    )


def _record_reference(record: Mapping[str, object]) -> dict[str, object]:
    return {
        "caseId": record["caseId"],
        "layer": record["layer"],
        **identity_core(record),
    }


def _normalized_fixture_diagnostics(fixture: Mapping[str, object]) -> list[dict[str, object]]:
    expected = fixture.get("expected")
    raw = expected.get("diagnostics") if isinstance(expected, dict) else None
    if not isinstance(raw, list):
        raise V4EvidenceContractError("source_fixture_diagnostics_invalid")
    normalized: list[dict[str, object]] = []
    for item in raw:
        if not isinstance(item, dict):
            raise V4EvidenceContractError("source_fixture_diagnostic_invalid")
        value = {
            "severity": item.get("severity"),
            "code": item.get("code"),
            "range": item.get("range"),
            "relatedRanges": item.get("related_ranges", []),
        }
        if not _diagnostic_shape(value):
            raise V4EvidenceContractError("source_fixture_diagnostic_shape_invalid")
        normalized.append(value)
    return normalized


def source_fixture_identities(record_values: Sequence[Mapping[str, object]]) -> list[dict[str, object]]:
    identities: dict[str, dict[str, object]] = {}
    for value in record_values:
        if value.get("layer") != "source_oracle":
            continue
        observation = cast(Mapping[str, object], value["observation"])
        payload = cast(Mapping[str, object], observation["payload"])
        fixture = cast(dict[str, object], payload["fixture"])
        relative = validate_relative_path(fixture.get("relativePath"), label="source_fixture")
        if relative in identities and identities[relative] != fixture:
            raise V4EvidenceContractError(f"source_fixture_identity_conflict:{relative}")
        identities[relative] = fixture
    return [identities[key] for key in sorted(identities)]


def evidence_artifact_identities(record_values: Sequence[Mapping[str, object]]) -> list[dict[str, object]]:
    fields = {
        "machine_wire": ("transcript",),
        "roundtrip": ("input", "output"),
        "headless_ooxml": ("artifact",),
        **dict.fromkeys(HOST_LAYERS, ("artifact",)),
    }
    identities: dict[str, dict[str, object]] = {}
    for value in record_values:
        layer = str(value.get("layer"))
        observation = cast(Mapping[str, object], value["observation"])
        payload = cast(Mapping[str, object], observation["payload"])
        selected = list(fields.get(layer, ()))
        if layer == "packaged":
            invocation = cast(Mapping[str, object], payload["invocation"])
            selected.extend(("invocation.stdout", "invocation.stderr"))
        for field in selected:
            identity = (
                cast(dict[str, object], invocation[field.split(".")[1]])
                if "." in field
                else cast(dict[str, object], payload[field])
            )
            relative = validate_relative_path(identity.get("relativePath"), label="evidence_artifact")
            if not relative.startswith("evidence/artifacts/"):
                raise V4EvidenceContractError(f"evidence_artifact_path_invalid:{relative}")
            if relative in identities and identities[relative] != identity:
                raise V4EvidenceContractError(f"evidence_artifact_identity_conflict:{relative}")
            identities[relative] = identity
    return [identities[key] for key in sorted(identities)]


def verify_observations(
    root: Path,
    records: Sequence[Mapping[str, object]],
    *,
    source_manifest_identity: Mapping[str, object],
    source_manifest: Mapping[str, object],
    package_identity: Mapping[str, object],
    package_manifest: Mapping[str, object],
) -> None:
    indexed = {(str(item["layer"]), str(item["caseId"])): item for item in records}
    values: dict[tuple[str, str], Mapping[str, object]] = {}
    for key, record in indexed.items():
        value, identity = read_json_object(
            root / str(record["relativePath"]), relative_to=root, label=f"observation:{key[0]}:{key[1]}"
        )
        if identity_core(identity) != identity_core(record):
            raise V4EvidenceContractError(f"observation_record_identity_mismatch:{key[0]}:{key[1]}")
        values[key] = value
    manifest_relative = validate_relative_path(source_manifest_identity.get("relativePath"), label="source_manifest")
    manifest_base = Path(manifest_relative).parent
    manifest_files = source_manifest.get("files")
    if not isinstance(manifest_files, list):
        raise V4EvidenceContractError("source_manifest_files_invalid")

    def payload(key: tuple[str, str]) -> Mapping[str, object]:
        observation = cast(Mapping[str, object], values[key]["observation"])
        return cast(Mapping[str, object], observation["payload"])

    def verify_artifact(identity: Mapping[str, object], label: str) -> None:
        relative = validate_relative_path(identity.get("relativePath"), label=label)
        if not relative.startswith("evidence/artifacts/"):
            raise V4EvidenceContractError(f"{label}_path_invalid:{relative}")
        if file_identity(root / relative, relative_to=root) != identity:
            raise V4EvidenceContractError(f"{label}_identity_mismatch")

    for key in indexed:
        layer, case_id = key
        data = payload(key)
        if layer == "source_oracle":
            fixture_identity = cast(Mapping[str, object], data["fixture"])
            relative = validate_relative_path(fixture_identity.get("relativePath"), label="source_fixture")
            try:
                oracle_relative = Path(relative).relative_to(manifest_base).as_posix()
            except ValueError as exc:
                raise V4EvidenceContractError(f"source_fixture_outside_oracle:{relative}") from exc
            matching = [
                item for item in manifest_files if isinstance(item, dict) and item.get("path") == oracle_relative
            ]
            if len(matching) != 1 or {
                "bytes": matching[0].get("bytes"),
                "sha256": matching[0].get("sha256"),
            } != {"bytes": fixture_identity.get("bytes"), "sha256": fixture_identity.get("sha256")}:
                raise V4EvidenceContractError(f"source_fixture_not_manifest_bound:{case_id}")
            fixture, actual = read_json_object(root / relative, relative_to=root, label=f"source_fixture:{case_id}")
            if identity_core(actual) != identity_core(fixture_identity) or fixture.get("case_id") != case_id:
                raise V4EvidenceContractError(f"source_fixture_identity_mismatch:{case_id}")
            source = fixture.get("source")
            if not isinstance(source, str) or hashlib.sha256(source.encode()).hexdigest() != data.get("sourceSha256"):
                raise V4EvidenceContractError(f"source_fixture_source_mismatch:{case_id}")
            if _normalized_fixture_diagnostics(fixture) != data.get("expectedDiagnostics"):
                raise V4EvidenceContractError(f"source_fixture_diagnostics_mismatch:{case_id}")
            expected_dimension = {
                value: dimension
                for dimension, value in {
                    "dot": "invalid-id-dot",
                    "slash": "invalid-id-slash",
                    "underscore": "invalid-id-underscore",
                    "over128": "invalid-id-too-long",
                    "duplicate": "invalid-duplicate",
                    "kindMismatch": "invalid-kind-mismatch",
                }.items()
            }.get(case_id)
            if data.get("invalidIdDimension") != expected_dimension:
                raise V4EvidenceContractError(f"source_fixture_invalid_dimension_mismatch:{case_id}")
        elif layer == "source_wire_comparison":
            source_ref = cast(Mapping[str, object], data["sourceRecord"])
            wire_ref = cast(Mapping[str, object], data["wireRecord"])
            source_key = ("source_oracle", str(source_ref["caseId"]))
            wire_key = ("machine_wire", str(wire_ref["caseId"]))
            if source_key not in indexed or source_ref != _record_reference(indexed[source_key]):
                raise V4EvidenceContractError(f"comparison_source_record_mismatch:{case_id}")
            if wire_key not in indexed or wire_ref != _record_reference(indexed[wire_key]):
                raise V4EvidenceContractError(f"comparison_wire_record_mismatch:{case_id}")
            source_data, wire_data = payload(source_key), payload(wire_key)
            terminal = cast(Mapping[str, object], wire_data["terminal"])
            params = cast(Mapping[str, object], terminal["params"])
            diagnostics = cast(list[Mapping[str, object]], params["diagnostics"])
            normalized = [
                {
                    "severity": item["severity"],
                    "code": item["code"],
                    "range": item["range"],
                    "relatedRanges": item["related_ranges"],
                }
                for item in diagnostics
            ]
            if normalized != source_data["expectedDiagnostics"] or any(
                cast(Mapping[str, object], item["source"])["sha256"] != source_data["sourceSha256"]
                for item in diagnostics
            ):
                raise V4EvidenceContractError(f"comparison_diagnostics_mismatch:{case_id}")
        elif layer == "machine_wire":
            transcript = cast(Mapping[str, object], data["transcript"])
            verify_artifact(transcript, f"machine_wire_transcript:{case_id}")
            transcript_value, _ = read_json_object(
                root / str(transcript["relativePath"]), relative_to=root, label=f"machine_wire_transcript:{case_id}"
            )
            if transcript_value != data["terminal"]:
                raise V4EvidenceContractError(f"machine_wire_transcript_mismatch:{case_id}")
        elif layer == "packaged":
            if data["packageManifest"] != identity_core(package_identity):
                raise V4EvidenceContractError(f"packaged_manifest_identity_mismatch:{case_id}")
            executable = cast(Mapping[str, object], data["executable"])
            actual = file_identity(root / str(executable["relativePath"]), relative_to=root)
            files = package_manifest.get("files")
            relative = Path(str(executable["relativePath"]))
            package_entry = {"package": relative.parts[0], "path": Path(*relative.parts[1:]).as_posix()}
            if (
                actual != executable
                or not isinstance(files, list)
                or not any(
                    isinstance(item, dict)
                    and all(item.get(name) == value for name, value in package_entry.items())
                    and item.get("bytes") == executable["bytes"]
                    and item.get("sha256") == executable["sha256"]
                    for item in files
                )
            ):
                raise V4EvidenceContractError(f"packaged_executable_identity_mismatch:{case_id}")
            invocation = cast(Mapping[str, object], data["invocation"])
            if cast(list[object], invocation["argv"])[0] != executable["relativePath"]:
                raise V4EvidenceContractError(f"packaged_argv_executable_mismatch:{case_id}")
            verify_artifact(cast(Mapping[str, object], invocation["stdout"]), f"packaged_stdout:{case_id}")
            verify_artifact(cast(Mapping[str, object], invocation["stderr"]), f"packaged_stderr:{case_id}")
        elif layer in {"roundtrip", "headless_ooxml", *HOST_LAYERS}:
            for field, expected_layer in (
                ("sourceRecord", "source_oracle"),
                ("packageRecord", "packaged"),
                ("headlessRecord", "headless_ooxml"),
            ):
                if field not in data:
                    continue
                reference = cast(Mapping[str, object], data[field])
                target = (expected_layer, str(reference["caseId"]))
                if target not in indexed or reference != _record_reference(indexed[target]):
                    raise V4EvidenceContractError(f"{layer}_{field}_mismatch:{case_id}")
            if layer == "roundtrip":
                input_identity = cast(Mapping[str, object], data["input"])
                output_identity = cast(Mapping[str, object], data["output"])
                verify_artifact(input_identity, f"roundtrip_input:{case_id}")
                verify_artifact(output_identity, f"roundtrip_output:{case_id}")
                source_reference = cast(Mapping[str, object], data["sourceRecord"])
                source_data = payload(("source_oracle", str(source_reference["caseId"])))
                if input_identity["sha256"] != source_data["sourceSha256"]:
                    raise V4EvidenceContractError(f"roundtrip_source_mismatch:{case_id}")
                if input_identity["sha256"] != output_identity["sha256"]:
                    raise V4EvidenceContractError(f"roundtrip_bytes_mismatch:{case_id}")
            else:
                verify_artifact(cast(Mapping[str, object], data["artifact"]), f"{layer}_artifact:{case_id}")


def manifest_expected(
    *,
    layer: str,
    records: Sequence[Mapping[str, object]],
    source_pointer: Mapping[str, object],
    wire_pointer: Mapping[str, object],
    package_pointer: Mapping[str, object],
) -> dict[str, object]:
    record_entries = sorted(
        ({"caseId": str(item["caseId"]), **identity_core(item)} for item in records),
        key=lambda item: str(item["caseId"]),
    )
    common: dict[str, object] = {"result": "passed", "records": record_entries}
    if layer == "source_wire_comparison":
        return {
            "schema": "docwen.v4_source_wire_comparison_manifest.v1",
            **common,
            "sourceOracle": dict(source_pointer),
            "machineWire": dict(wire_pointer),
        }
    if layer in {"roundtrip", "headless_ooxml"}:
        schema = "docwen.v4_roundtrip_manifest.v1" if layer == "roundtrip" else "docwen.v4_headless_ooxml_manifest.v1"
        return {"schema": schema, **common, "packageManifest": dict(package_pointer)}
    if layer in HOST_LAYERS:
        return {
            "schema": "docwen.v4_host_manifest.v1",
            "host": layer.removesuffix("_host"),
            **common,
            "packageManifest": dict(package_pointer),
        }
    raise ValueError(f"external_manifest_layer_invalid:{layer}")
