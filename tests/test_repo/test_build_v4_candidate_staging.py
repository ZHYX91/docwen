from __future__ import annotations

import json
import shutil
from collections.abc import Mapping
from pathlib import Path
from typing import Any, cast

import pytest
from jsonschema import Draft202012Validator
from scripts.release import build_v4_candidate_staging as builder
from scripts.release import candidate_evidence
from scripts.release import v4_candidate_contract as contract
from scripts.release import v4_evidence_contract as evidence_contract
from tests.test_repo import v4_evidence_test_support as support

pytestmark = pytest.mark.contract
ROOT = Path(__file__).resolve().parents[2]
DOCWEN_COMMIT = "0123456789abcdef0123456789abcdef01234567"
CANDIDATE_ID = f"docwen-0.9.0-v4-20260814T120000Z-{DOCWEN_COMMIT[:12]}"


def _write_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(contract.json_bytes(value))


def _copy_file(source: Path, root: Path, relative: str) -> dict[str, object]:
    destination = root / relative
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, destination)
    return cast(dict[str, object], contract.file_identity(destination, relative_to=root))


def _synthetic_source(tmp_path: Path) -> tuple[Path, dict[str, object]]:
    docwen = tmp_path / "docwen"
    docwen.mkdir()
    authority = _copy_file(ROOT / contract.AUTHORITY_PATH, docwen, contract.AUTHORITY_PATH)
    authority_value = contract.read_json_object(docwen / contract.AUTHORITY_PATH, label="synthetic_authority")
    oracle_relative = "contracts/oracles/docwen.markdown_semantics.v3/manifest.json"
    _copy_file(ROOT / oracle_relative, docwen, oracle_relative)
    oracle_value = contract.read_json_object(docwen / oracle_relative, label="synthetic_oracle")
    oracle_value["final_spec_baseline"] = authority_value["docwenSpecBaseline"]
    _write_json(docwen / oracle_relative, oracle_value)
    oracle = contract.file_identity(docwen / oracle_relative, relative_to=docwen)
    for case_id in contract.REQUIRED_INVALID_CASES.values():
        relative = f"contracts/oracles/docwen.markdown_semantics.v3/corpus/{case_id}.case.json"
        _copy_file(ROOT / relative, docwen, relative)
    spec_files = [
        _copy_file(ROOT / relative, docwen, relative)
        for relative in (contract.RECEIPT_SCHEMA_PATH, contract.EVIDENCE_SCHEMA_PATH, contract.WIRE_SCHEMA_PATH)
    ]
    identity = {"commit": DOCWEN_COMMIT, "tree": "1" * 40}
    source = {
        "schema": "docwen.v4_candidate_source_checkpoint.v1",
        "authorityBaseline": authority,
        "activeSemantics": authority_value["activeSemantics"],
        "finalSpecIdentity": {"files": spec_files, "sha256": contract.payload_sha256(spec_files)},
        "oracleManifest": oracle,
        "docwen": {"specBaseline": identity, "implementationBaseline": identity, "final": identity},
        "exclusions": authority_value["exclusions"],
        "ignoredExecutableInputs": [],
        "ignoredSourceInputs": [],
    }
    return docwen, source


def _source_staging(tmp_path: Path) -> Path:
    root = tmp_path / "package-input"
    gui = root / builder.PACKAGE_NAMES[0]
    cli = root / builder.PACKAGE_NAMES[1]
    gui.mkdir(parents=True)
    cli.mkdir()
    (gui / "DocWen.exe").write_bytes(b"gui")
    (gui / "DocWenCLI.exe").write_bytes(b"cli")
    (cli / "DocWenCLI.exe").write_bytes(b"cli")
    return root


def _record(path: Path, *, case_id: str, layer: str, payload: Mapping[str, object]) -> None:
    _write_json(
        path,
        {
            "schema": builder.RECORD_SCHEMA,
            "caseId": case_id,
            "layer": layer,
            "result": "passed",
            "observation": {"kind": layer, "payload": dict(payload)},
        },
    )


def _fixture(tmp_path: Path, *, hosts: str = "not_run") -> tuple[dict[str, Any], dict[str, object]]:
    docwen, source = _synthetic_source(tmp_path)
    package_input = _source_staging(tmp_path)
    evidence = tmp_path / "evidence-input"
    evidence.mkdir()
    records: list[dict[str, str]] = []
    predicted: list[dict[str, object]] = []

    def add(case_id: str, layer: str, payload: Mapping[str, object]) -> dict[str, object]:
        relative = f"records/{layer}/{case_id}.json"
        path = evidence / relative
        _record(path, case_id=case_id, layer=layer, payload=payload)
        source_identity = contract.file_identity(path, relative_to=evidence)
        records.append({"caseId": case_id, "layer": layer, "artifact": relative})
        value = {
            "caseId": case_id,
            "layer": layer,
            **builder._identity_at(source_identity, builder._record_destination(layer, case_id)),
        }
        predicted.append(value)
        return value

    source_records: dict[str, tuple[dict[str, object], dict[str, object]]] = {}
    for dimension, case_id in contract.REQUIRED_INVALID_CASES.items():
        payload = support.source_observation(docwen, case_id, dimension)
        source_records[case_id] = (payload, add(case_id, "source_oracle", payload))
    dot_payload, dot_record = source_records["invalid-id-dot"]
    terminal = support.wire_terminal(dot_payload)
    transcript = support.evidence_artifact(evidence, "machine-terminal.json", terminal)
    wire_payload = {
        "schema": "docwen.v4_machine_wire_observation.v1",
        "protocol": "docwen.machine.v1",
        "transcript": transcript,
        "terminal": terminal,
        "terminalSha256": evidence_contract._payload_hash(terminal),
    }
    wire_record = add("positive-machine-wire", "machine_wire", wire_payload)
    package_payload = candidate_evidence.capture_package_manifest(package_input, builder.PACKAGE_NAMES)
    package_pointer = builder._with_role(
        "package_manifest",
        builder._json_identity(package_payload, relative_path=contract.PACKAGE_MANIFEST_PATH),
    )
    executable = contract.file_identity(
        package_input / builder.PACKAGE_NAMES[1] / "DocWenCLI.exe", relative_to=package_input
    )
    stdout = support.evidence_artifact(evidence, "packaged-stdout.bin", b"")
    stderr = support.evidence_artifact(evidence, "packaged-stderr.bin", b"")
    package_record = add(
        "positive-packaged",
        "packaged",
        {
            "schema": "docwen.v4_packaged_observation.v1",
            "packageManifest": evidence_contract.identity_core(package_pointer),
            "executable": executable,
            "invocation": {
                "argv": [executable["relativePath"], "--version"],
                "exitCode": 0,
                "stdout": stdout,
                "stderr": stderr,
            },
        },
    )
    add(
        "positive-comparison",
        "source_wire_comparison",
        {
            "schema": "docwen.v4_source_wire_comparison_observation.v1",
            "result": "equal",
            "sourceRecord": support.record_ref(dot_record),
            "wireRecord": support.record_ref(wire_record),
            "comparedFields": ["diagnostics"],
            "mismatches": [],
        },
    )
    roundtrip_bytes = support.source_bytes(docwen, "invalid-id-dot")
    roundtrip_input = support.evidence_artifact(evidence, "roundtrip-input.md", roundtrip_bytes)
    roundtrip_output = support.evidence_artifact(evidence, "roundtrip-output.md", roundtrip_bytes)
    add(
        "positive-roundtrip",
        "roundtrip",
        {
            "schema": "docwen.v4_roundtrip_observation.v1",
            "sourceRecord": support.record_ref(dot_record),
            "packageRecord": support.record_ref(package_record),
            "input": roundtrip_input,
            "output": roundtrip_output,
            "byteExact": True,
        },
    )
    headless_artifact = support.evidence_artifact(evidence, "headless.docx", b"headless-ooxml")
    headless_record = add(
        "positive-headless",
        "headless_ooxml",
        {
            "schema": "docwen.v4_headless_ooxml_observation.v1",
            "packageRecord": support.record_ref(package_record),
            "artifact": headless_artifact,
            "inspection": {"bookmarkCount": 1, "seqFieldCount": 1, "refFieldCount": 1, "violations": []},
        },
    )
    if hosts == "passed":
        for layer in contract.REQUIRED_LAYERS[6:]:
            host_artifact = support.evidence_artifact(evidence, f"{layer}.docx", layer.encode())
            add(
                f"positive-{layer.replace('_', '-')}",
                layer,
                {
                    "schema": "docwen.v4_host_observation.v1",
                    "packageRecord": support.record_ref(package_record),
                    "headlessRecord": support.record_ref(headless_record),
                    "artifact": host_artifact,
                    "host": {
                        "name": layer.removesuffix("_host"),
                        "version": "synthetic-1",
                        "opened": True,
                        "rendered": True,
                        "saved": True,
                        "violations": [],
                    },
                },
            )
    oracle = cast(Mapping[str, object], source["oracleManifest"])
    final_files = cast(
        Mapping[str, Mapping[str, object]],
        {
            item["relativePath"]: item
            for item in cast(list[dict[str, object]], cast(Mapping[str, object], source["finalSpecIdentity"])["files"])
        },
    )
    source_pointer = builder._with_role("source_manifest", oracle)
    wire_pointer = builder._with_role("wire_schema", final_files[contract.WIRE_SCHEMA_PATH])
    manifests: dict[str, str | None] = {}
    for layer in builder.EXTERNAL_MANIFEST_LAYERS:
        layer_records = [item for item in predicted if item["layer"] == layer]
        if not layer_records:
            manifests[layer] = None
            continue
        relative = f"manifests/{layer}.json"
        _write_json(
            evidence / relative,
            builder._manifest_expected(
                layer=layer,
                records=layer_records,
                source_pointer=source_pointer,
                wire_pointer=wire_pointer,
                package_pointer=package_pointer,
            ),
        )
        manifests[layer] = relative
    statuses = {
        **dict.fromkeys(contract.REQUIRED_LAYERS[:6], "passed"),
        **dict.fromkeys(contract.REQUIRED_LAYERS[6:], hosts),
    }
    external = tmp_path / "external"
    external.mkdir()
    source_path = external / "source.json"
    _write_json(source_path, source)
    plan = {
        "schema": builder.PLAN_SCHEMA,
        "candidateId": CANDIDATE_ID,
        "layerStatus": statuses,
        "records": records,
        "manifests": manifests,
    }
    plan_path = external / "plan.json"
    _write_json(plan_path, plan)
    output = tmp_path / "output"
    output.mkdir()
    args: dict[str, Any] = {
        "docwen_repo": docwen,
        "source_staging": package_input,
        "evidence_root": evidence,
        "evidence_plan": plan_path,
        "evidence_plan_sha256": contract.sha256_file(plan_path),
        "source_checkpoint": source_path,
        "source_checkpoint_sha256": contract.sha256_file(source_path),
        "staging_checkpoint_output": external / "staging.json",
        "output_root": output,
        "candidate_id": CANDIDATE_ID,
    }
    return args, source


def test_builds_closed_layered_index_and_external_staging_checkpoint(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    args, source = _fixture(tmp_path)
    monkeypatch.setattr(contract, "capture_source_checkpoint", lambda **_kwargs: source)
    result = builder.build_staging(**args)
    staging = Path(cast(str, result["stagingRoot"]))
    index = contract.read_json_object(staging / "_evidence/v4-evidence-index.json", label="index")
    records, _ = contract.validate_evidence_index(index, candidate_id=CANDIDATE_ID)
    assert not [item for item in records if str(item["layer"]).endswith("_host")]
    assert {cast(list[dict[str, object]], item["identityPointers"])[0]["role"] for item in records} == {
        "source_manifest",
        "wire_schema",
        "comparison_manifest",
        "package_manifest",
        "roundtrip_manifest",
        "headless_ooxml_manifest",
    }
    assert not any(
        cast(list[dict[str, object]], item["identityPointers"])[0]["relativePath"]
        in {"receipt.json", "candidate.json", "manifest.json", "_evidence/v4-evidence-index.json"}
        for item in records
    )
    for dimension, case_id in contract.REQUIRED_INVALID_CASES.items():
        record = next(item for item in records if item["caseId"] == case_id and item["layer"] == "source_oracle")
        value = contract.read_json_object(staging / cast(str, record["relativePath"]), label=case_id)
        payload = cast(Mapping[str, object], cast(Mapping[str, object], value["observation"])["payload"])
        fixture_identity = cast(Mapping[str, object], payload["fixture"])
        fixture = contract.read_json_object(
            staging / cast(str, fixture_identity["relativePath"]), label=f"fixture:{case_id}"
        )
        assert payload["invalidIdDimension"] == dimension
        assert fixture_identity == contract.file_identity(
            staging / cast(str, fixture_identity["relativePath"]), relative_to=staging
        )
        assert payload["expectedDiagnostics"] == evidence_contract._normalized_fixture_diagnostics(fixture)
    checkpoint = contract.read_json_object(cast(Path, args["staging_checkpoint_output"]), label="checkpoint")
    assert checkpoint["sourceCheckpointPayloadSha256"] == contract.payload_sha256(source)
    assert result["sourceCheckpointSha256"] == args["source_checkpoint_sha256"]
    assert result["evidencePlan"]["sha256"] == args["evidence_plan_sha256"]
    assert candidate_evidence._capture_tree_stable(staging)["manifestSha256"]


def test_all_passed_hosts_have_exactly_one_closed_manifest(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    args, source = _fixture(tmp_path, hosts="passed")
    monkeypatch.setattr(contract, "capture_source_checkpoint", lambda **_kwargs: source)
    staging = Path(cast(str, builder.build_staging(**args)["stagingRoot"]))
    index = contract.read_json_object(staging / "_evidence/v4-evidence-index.json", label="index")
    records, _ = contract.validate_evidence_index(index, candidate_id=CANDIDATE_ID)
    for layer in contract.REQUIRED_LAYERS[6:]:
        matches = [item for item in records if item["layer"] == layer]
        assert len(matches) == 1
        manifest = contract.read_json_object(staging / builder._manifest_destination(layer), label=layer)
        assert set(manifest) == {"schema", "host", "result", "records", "packageManifest"}
        assert manifest["host"] == layer.removesuffix("_host")


def test_plan_or_manifest_claim_cannot_substitute_for_closed_evidence(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    args, source = _fixture(tmp_path)
    monkeypatch.setattr(contract, "capture_source_checkpoint", lambda **_kwargs: source)
    evidence = cast(Path, args["evidence_root"])
    record = next(evidence.glob("records/machine_wire/*.json"))
    value = contract.read_json_object(record, label="record")
    cast(dict[str, object], value["observation"])["payload"] = {"synthetic": True}
    _write_json(record, value)
    with pytest.raises(builder.V4StagingBuildError, match="evidence_record_envelope_mismatch"):
        builder.build_staging(**args)
    assert not list(cast(Path, args["output_root"]).iterdir())


def test_comparison_must_bind_exact_indexed_source_and_wire_records(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    args, source = _fixture(tmp_path)
    monkeypatch.setattr(contract, "capture_source_checkpoint", lambda **_kwargs: source)
    staging = Path(cast(str, builder.build_staging(**args)["stagingRoot"]))
    index = contract.read_json_object(staging / "_evidence/v4-evidence-index.json", label="index")
    records = cast(list[dict[str, object]], index["records"])
    comparison_record = next(item for item in records if item["layer"] == "source_wire_comparison")
    comparison = staging / cast(str, comparison_record["relativePath"])
    value = contract.read_json_object(comparison, label="comparison")
    payload = cast(dict[str, object], cast(dict[str, object], value["observation"])["payload"])
    cast(dict[str, object], payload["sourceRecord"])["sha256"] = "f" * 64
    _write_json(comparison, value)
    comparison_record.update(contract.file_identity(comparison, relative_to=staging))
    authority, _ = contract.load_authority(staging)
    source_identity = contract.oracle_manifest_identity(staging, authority)
    package_identity = contract.file_identity(staging / contract.PACKAGE_MANIFEST_PATH, relative_to=staging)
    with pytest.raises(evidence_contract.V4EvidenceContractError, match="comparison_source_record_mismatch"):
        evidence_contract.verify_observations(
            staging,
            records,
            source_manifest_identity=source_identity,
            source_manifest=contract.read_json_object(
                staging / cast(str, source_identity["relativePath"]), label="source_manifest"
            ),
            package_identity=package_identity,
            package_manifest=contract.read_json_object(
                staging / contract.PACKAGE_MANIFEST_PATH, label="package_manifest"
            ),
        )


def test_external_json_duplicate_keys_fail_before_staging_write(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    args, source = _fixture(tmp_path)
    monkeypatch.setattr(contract, "capture_source_checkpoint", lambda **_kwargs: source)
    plan = cast(Path, args["evidence_plan"])
    plan.write_bytes(plan.read_bytes().replace(b"{", b'{"schema":"duplicate",', 1))
    args["evidence_plan_sha256"] = contract.sha256_file(plan)
    with pytest.raises(builder.V4StagingBuildError, match="duplicate_json_key:schema"):
        builder.build_staging(**args)
    assert not list(cast(Path, args["output_root"]).iterdir())


def test_external_plan_mutation_before_publish_is_quarantined(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    args, source = _fixture(tmp_path)
    monkeypatch.setattr(contract, "capture_source_checkpoint", lambda **_kwargs: source)
    original = builder._copy_checked
    changed = False

    def mutate_plan(*call_args: object, **call_kwargs: object) -> dict[str, object]:
        nonlocal changed
        result = original(*call_args, **call_kwargs)
        if not changed:
            changed = True
            cast(Path, args["evidence_plan"]).write_bytes(cast(Path, args["evidence_plan"]).read_bytes() + b" ")
        return result

    monkeypatch.setattr(builder, "_copy_checked", mutate_plan)
    with pytest.raises(builder.V4StagingBuildError, match="evidence_plan_sha256_mismatch"):
        builder.build_staging(**args)
    assert not (cast(Path, args["output_root"]) / f"{CANDIDATE_ID}-staging").exists()


def test_path_and_tree_safety_rejects_colon_hardlink_and_duplicate_inode(tmp_path: Path) -> None:
    with pytest.raises(builder.V4StagingBuildError, match="artifact_path_invalid"):
        builder._safe_relative("record.json:stream", label="artifact")
    outside = tmp_path / "outside.bin"
    outside.write_bytes(b"same")
    one = tmp_path / "one"
    one.mkdir()
    (one / "a.bin").hardlink_to(outside)
    with pytest.raises(contract.V4CandidateContractError, match="tree_hardlink_rejected"):
        contract.capture_tree_stable(one)
    two = tmp_path / "two"
    two.mkdir()
    (two / "a.bin").write_bytes(b"same")
    (two / "b.bin").hardlink_to(two / "a.bin")
    with pytest.raises(contract.V4CandidateContractError, match="duplicate_tree_inode"):
        contract.capture_tree_stable(two)


def test_pending_source_fails_before_any_staging_write(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    args, _ = _fixture(tmp_path)

    def pending(**_kwargs: object) -> dict[str, object]:
        raise contract.V4CandidateContractError("docwen_implementation_baseline_pending")

    monkeypatch.setattr(contract, "capture_source_checkpoint", pending)
    with pytest.raises(contract.V4CandidateContractError, match="implementation_baseline_pending"):
        builder.build_staging(**args)
    assert not list(cast(Path, args["output_root"]).iterdir())
    assert not cast(Path, args["staging_checkpoint_output"]).exists()


def test_authority_byte_drift_fails_before_any_staging_write(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    args, source = _fixture(tmp_path)
    monkeypatch.setattr(contract, "capture_source_checkpoint", lambda **_kwargs: source)
    authority = cast(Path, args["docwen_repo"]) / contract.AUTHORITY_PATH
    authority.write_bytes(authority.read_bytes() + b" ")
    with pytest.raises(builder.V4StagingBuildError, match="source_governance_bytes_drifted"):
        builder.build_staging(**args)
    assert not list(cast(Path, args["output_root"]).iterdir())
    assert not cast(Path, args["staging_checkpoint_output"]).exists()


def test_partial_copy_is_quarantined_and_never_published(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    args, source = _fixture(tmp_path)
    monkeypatch.setattr(contract, "capture_source_checkpoint", lambda **_kwargs: source)
    original = builder._copy_checked
    calls = 0

    def fail_after_one(*call_args: object, **call_kwargs: object) -> dict[str, object]:
        nonlocal calls
        calls += 1
        result = original(*call_args, **call_kwargs)
        if calls == 2:
            raise builder.V4StagingBuildError("synthetic_copy_toctou")
        return result

    monkeypatch.setattr(builder, "_copy_checked", fail_after_one)
    with pytest.raises(builder.V4StagingBuildError, match="synthetic_copy_toctou"):
        builder.build_staging(**args)
    output = cast(Path, args["output_root"])
    assert not (output / f"{CANDIDATE_ID}-staging").exists()
    assert len(list(output.glob(f".rejected-{CANDIDATE_ID}-*"))) == 1
    assert not cast(Path, args["staging_checkpoint_output"]).exists()


def test_post_checkpoint_failure_quarantines_both_outputs(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    args, source = _fixture(tmp_path)
    monkeypatch.setattr(contract, "capture_source_checkpoint", lambda **_kwargs: source)
    original = builder.sealer.write_staging_checkpoint

    def fail_after_checkpoint(**kwargs: object) -> dict[str, object]:
        original(**kwargs)
        raise builder.V4StagingBuildError("synthetic_post_checkpoint_failure")

    monkeypatch.setattr(builder.sealer, "write_staging_checkpoint", fail_after_checkpoint)
    with pytest.raises(builder.V4StagingBuildError, match="synthetic_post_checkpoint_failure"):
        builder.build_staging(**args)
    output = cast(Path, args["output_root"])
    checkpoint = cast(Path, args["staging_checkpoint_output"])
    assert not (output / f"{CANDIDATE_ID}-staging").exists()
    assert not checkpoint.exists()
    assert len(list(output.glob(f".rejected-{CANDIDATE_ID}-*"))) == 1
    assert len(list(checkpoint.parent.glob(f".rejected-{checkpoint.name}-*"))) == 1


def test_quarantine_rename_failure_is_combined_and_canonical_is_nonreusable(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    args, source = _fixture(tmp_path)
    monkeypatch.setattr(contract, "capture_source_checkpoint", lambda **_kwargs: source)
    original_write = builder.sealer.write_staging_checkpoint
    original_rename = Path.rename

    def fail_after_checkpoint(**kwargs: object) -> dict[str, object]:
        original_write(**kwargs)
        raise builder.V4StagingBuildError("synthetic_proof_failure")

    def refuse_quarantine(path: Path, target: Path) -> Path:
        if target.name.startswith(".rejected-"):
            raise OSError("synthetic_quarantine_failure")
        return original_rename(path, target)

    monkeypatch.setattr(builder.sealer, "write_staging_checkpoint", fail_after_checkpoint)
    monkeypatch.setattr(Path, "rename", refuse_quarantine)
    with pytest.raises(builder.V4StagingBuildError, match="staging_failure_and_quarantine_failed"):
        builder.build_staging(**args)
    canonical = cast(Path, args["output_root"]) / f"{CANDIDATE_ID}-staging"
    assert canonical.exists()
    with pytest.raises(builder.V4StagingBuildError, match="final_staging_path_exists"):
        builder.build_staging(**args)


def test_verifier_rejects_opaque_record_even_after_rehash(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    args, source = _fixture(tmp_path)
    monkeypatch.setattr(contract, "capture_source_checkpoint", lambda **_kwargs: source)
    staging = Path(cast(str, builder.build_staging(**args)["stagingRoot"]))
    index_path = staging / "_evidence/v4-evidence-index.json"
    index = contract.read_json_object(index_path, label="index")
    records = cast(list[dict[str, object]], index["records"])
    record = next(item for item in records if item["caseId"] == "invalid-id-dot")
    evidence_path = staging / cast(str, record["relativePath"])
    _write_json(evidence_path, {"claim": "passed"})
    record.update(contract.file_identity(evidence_path, relative_to=staging))
    _write_json(index_path, index)
    with pytest.raises(contract.V4CandidateContractError, match="evidence_record_envelope_mismatch"):
        contract.verify_index_files(staging, records, stable_tree=candidate_evidence._capture_tree_stable(staging))


def test_source_checkpoint_verification_recomputes_authority_payload(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    args, source = _fixture(tmp_path)
    monkeypatch.setattr(contract, "capture_source_checkpoint", lambda **_kwargs: source)
    drifted = json.loads(json.dumps(source))
    cast(dict[str, object], drifted["authorityBaseline"])["sha256"] = "f" * 64
    with pytest.raises(contract.V4CandidateContractError, match="source_checkpoint_payload_mismatch"):
        contract.verify_source_checkpoint(
            drifted,
            docwen_repo=cast(Path, args["docwen_repo"]),
        )


def test_evidence_schema_rejects_not_run_host_record_and_cross_layer_role(tmp_path: Path) -> None:
    args, _ = _fixture(tmp_path)
    plan = contract.read_json_object(cast(Path, args["evidence_plan"]), label="plan")
    records = cast(list[dict[str, object]], plan["records"])
    source_record = records[0]
    synthetic_index = {
        "schema": contract.EVIDENCE_SCHEMA,
        "candidateId": CANDIDATE_ID,
        "layerStatus": plan["layerStatus"],
        "records": [
            {
                "caseId": source_record["caseId"],
                "layer": "source_oracle",
                "relativePath": "evidence/record.json",
                "bytes": 1,
                "sha256": "0" * 64,
                "identityPointers": [
                    {
                        "role": "wire_schema",
                        "relativePath": contract.WIRE_SCHEMA_PATH,
                        "bytes": 1,
                        "sha256": "1" * 64,
                    }
                ],
            },
            {
                "caseId": "unexpected-word",
                "layer": "word_host",
                "relativePath": "evidence/host.json",
                "bytes": 1,
                "sha256": "2" * 64,
                "identityPointers": [
                    {
                        "role": "host_manifest",
                        "relativePath": "evidence/host-manifest.json",
                        "bytes": 1,
                        "sha256": "3" * 64,
                    }
                ],
            },
        ],
    }
    schema = json.loads((ROOT / contract.EVIDENCE_SCHEMA_PATH).read_text(encoding="utf-8"))
    messages = "\n".join(error.message for error in Draft202012Validator(schema).iter_errors(synthetic_index))
    assert "source_manifest" in messages
    assert "does not contain items" in messages or "should not be valid" in messages


def test_v4_staging_tool_files_remain_under_governance_threshold() -> None:
    for relative in (
        "scripts/release/build_v4_candidate_staging.py",
        "scripts/release/v4_evidence_contract.py",
        "tests/test_repo/test_build_v4_candidate_staging.py",
    ):
        assert len((ROOT / relative).read_text(encoding="utf-8").splitlines()) < 700
