from __future__ import annotations

import json
import shutil
import subprocess
from collections.abc import Mapping
from pathlib import Path
from typing import Any, cast

import pytest
from jsonschema import Draft202012Validator
from scripts.release import build_v4_candidate_staging as staging_builder
from scripts.release import candidate_evidence
from scripts.release import seal_v4_candidate as sealer
from scripts.release import v4_candidate_contract as contract
from scripts.release import v4_evidence_contract as evidence_contract
from tests.test_repo import v4_evidence_test_support as support

pytestmark = pytest.mark.contract
ROOT = Path(__file__).resolve().parents[2]


def _git(repo: Path, *arguments: str) -> str:
    return subprocess.run(
        ["git", "-C", str(repo), *arguments],
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
    ).stdout.strip()


def _init_repo(path: Path, *, commits: int = 1) -> tuple[dict[str, str], dict[str, str]]:
    path.mkdir()
    _git(path, "init", "-b", "main")
    _git(path, "config", "user.name", "Candidate Test")
    _git(path, "config", "user.email", "candidate@example.invalid")
    baselines: list[dict[str, str]] = []
    for index in range(commits):
        (path / "source.txt").write_text(f"source-{index}\n", encoding="utf-8")
        _git(path, "add", "source.txt")
        _git(path, "commit", "-m", f"source {index}")
        baselines.append({"commit": _git(path, "rev-parse", "HEAD"), "tree": _git(path, "rev-parse", "HEAD^{tree}")})
    return baselines[0], baselines[-1]


def _identity(path: Path, root: Path, *, role: str | None = None) -> dict[str, object]:
    value: dict[str, object] = contract.file_identity(path, relative_to=root)
    if role is not None:
        value = {"role": role, **value}
    return value


def _write_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(contract.json_bytes(value))


def _staging(tmp_path: Path, candidate_id: str, *, host_status: str = "not_run") -> Path:
    root = tmp_path / "staging"
    gui = root / f"DocWen_v{contract.PRODUCT_VERSION}_win-x64"
    cli = root / f"DocWenCLI_v{contract.PRODUCT_VERSION}_win-x64"
    gui.mkdir(parents=True)
    cli.mkdir()
    (gui / "DocWen.exe").write_bytes(b"gui")
    (gui / "DocWenCLI.exe").write_bytes(b"cli")
    (cli / "DocWenCLI.exe").write_bytes(b"cli")
    for source in (
        contract.AUTHORITY_PATH,
        contract.RECEIPT_SCHEMA_PATH,
        contract.EVIDENCE_SCHEMA_PATH,
        contract.WIRE_SCHEMA_PATH,
        "contracts/oracles/docwen.markdown_semantics.v3/manifest.json",
    ):
        destination = root / source
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(ROOT / source, destination)
    staged_authority = contract.read_json_object(root / contract.AUTHORITY_PATH, label="synthetic_authority")
    staged_oracle_path = root / "contracts/oracles/docwen.markdown_semantics.v3/manifest.json"
    staged_oracle = contract.read_json_object(staged_oracle_path, label="synthetic_oracle")
    staged_oracle["final_spec_baseline"] = staged_authority["docwenSpecBaseline"]
    _write_json(staged_oracle_path, staged_oracle)
    for case_id in contract.REQUIRED_INVALID_CASES.values():
        relative = f"contracts/oracles/docwen.markdown_semantics.v3/corpus/{case_id}.case.json"
        destination = root / relative
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(ROOT / relative, destination)
    evidence_root = root / "_evidence"
    evidence_root.mkdir()
    candidate_evidence.write_package_manifest(
        root,
        [gui.name, cli.name],
        evidence_root / "package-manifest.json",
        allowed_root_entries={"_evidence", "contracts"},
    )
    package_identity = _identity(evidence_root / "package-manifest.json", root, role="package_manifest")
    records: list[dict[str, object]] = []

    def add(case_id: str, layer: str, payload: Mapping[str, object]) -> dict[str, object]:
        path = root / "evidence" / "records" / layer / f"{case_id}.json"
        _write_json(
            path,
            {
                "schema": staging_builder.RECORD_SCHEMA,
                "caseId": case_id,
                "layer": layer,
                "result": "passed",
                "observation": {"kind": layer, "payload": dict(payload)},
            },
        )
        record = {"caseId": case_id, "layer": layer, **_identity(path, root)}
        records.append(record)
        return record

    source_records: dict[str, tuple[dict[str, object], dict[str, object]]] = {}
    for dimension, case_id in contract.REQUIRED_INVALID_CASES.items():
        payload = support.source_observation(root, case_id, dimension)
        source_records[case_id] = (payload, add(case_id, "source_oracle", payload))
    dot_payload, dot_record = source_records["invalid-id-dot"]
    terminal = support.wire_terminal(dot_payload)
    transcript = support.evidence_artifact(root / "evidence", "machine-terminal.json", terminal)
    wire_record = add(
        "positive-machine-wire",
        "machine_wire",
        {
            "schema": "docwen.v4_machine_wire_observation.v1",
            "protocol": "docwen.machine.v1",
            "transcript": transcript,
            "terminal": terminal,
            "terminalSha256": evidence_contract._payload_hash(terminal),
        },
    )
    executable = _identity(cli / "DocWenCLI.exe", root)
    stdout = support.evidence_artifact(root / "evidence", "packaged-stdout.bin", b"")
    stderr = support.evidence_artifact(root / "evidence", "packaged-stderr.bin", b"")
    package_record = add(
        "positive-packaged",
        "packaged",
        {
            "schema": "docwen.v4_packaged_observation.v1",
            "packageManifest": evidence_contract.identity_core(package_identity),
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
        "positive-source-wire-comparison",
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
    roundtrip_bytes = support.source_bytes(root, "invalid-id-dot")
    roundtrip_input = support.evidence_artifact(root / "evidence", "roundtrip-input.md", roundtrip_bytes)
    roundtrip_output = support.evidence_artifact(root / "evidence", "roundtrip-output.md", roundtrip_bytes)
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
    headless_artifact = support.evidence_artifact(root / "evidence", "headless.docx", b"headless-ooxml")
    headless_record = add(
        "positive-headless-ooxml",
        "headless_ooxml",
        {
            "schema": "docwen.v4_headless_ooxml_observation.v1",
            "packageRecord": support.record_ref(package_record),
            "artifact": headless_artifact,
            "inspection": {"bookmarkCount": 1, "seqFieldCount": 1, "refFieldCount": 1, "violations": []},
        },
    )
    if host_status == "passed":
        for layer in contract.REQUIRED_LAYERS[6:]:
            host_artifact = support.evidence_artifact(root / "evidence", f"{layer}.docx", layer.encode())
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
    source_pointer = _identity(
        root / "contracts/oracles/docwen.markdown_semantics.v3/manifest.json", root, role="source_manifest"
    )
    wire_pointer = _identity(root / contract.WIRE_SCHEMA_PATH, root, role="wire_schema")
    pointers: dict[str, dict[str, object]] = {
        "source_oracle": source_pointer,
        "machine_wire": wire_pointer,
        "packaged": package_identity,
    }
    for layer in ("source_wire_comparison", "roundtrip", "headless_ooxml", *contract.REQUIRED_LAYERS[6:]):
        layer_records = [item for item in records if item["layer"] == layer]
        if not layer_records:
            continue
        manifest = root / staging_builder._manifest_destination(layer)
        _write_json(
            manifest,
            staging_builder._manifest_expected(
                layer=layer,
                records=layer_records,
                source_pointer=source_pointer,
                wire_pointer=wire_pointer,
                package_pointer=package_identity,
            ),
        )
        pointers[layer] = _identity(manifest, root, role=contract.POINTER_ROLE_BY_LAYER[layer])
    for record in records:
        record["identityPointers"] = [pointers[cast(str, record["layer"])]]
    index = {
        "schema": contract.EVIDENCE_SCHEMA,
        "candidateId": candidate_id,
        "layerStatus": {
            **dict.fromkeys(contract.REQUIRED_LAYERS[:6], "passed"),
            **dict.fromkeys(contract.REQUIRED_LAYERS[6:], host_status),
        },
        "records": records,
    }
    _write_json(evidence_root / "v4-evidence-index.json", index)
    return root


def _source_payload(docwen: Mapping[str, str], staging: Path) -> dict[str, object]:
    authority = contract.read_json_object(staging / contract.AUTHORITY_PATH, label="source_authority")
    spec_files = [
        _identity(staging / relative, staging)
        for relative in (contract.RECEIPT_SCHEMA_PATH, contract.EVIDENCE_SCHEMA_PATH, contract.WIRE_SCHEMA_PATH)
    ]
    return {
        "schema": sealer.SOURCE_CHECKPOINT_SCHEMA,
        "authorityBaseline": _identity(staging / contract.AUTHORITY_PATH, staging),
        "activeSemantics": authority["activeSemantics"],
        "finalSpecIdentity": {"files": spec_files, "sha256": contract.payload_sha256(spec_files)},
        "oracleManifest": _identity(staging / "contracts/oracles/docwen.markdown_semantics.v3/manifest.json", staging),
        "docwen": {
            "specBaseline": docwen,
            "implementationBaseline": docwen,
            "final": docwen,
        },
        "exclusions": authority["exclusions"],
        "ignoredExecutableInputs": [],
        "ignoredSourceInputs": [],
    }


def _checkpoints(tmp_path: Path, staging: Path, source: Mapping[str, object]) -> tuple[Path, str, Path, str]:
    source_path = tmp_path / "external" / "source.json"
    _write_json(source_path, source)
    source_sha = contract.sha256_file(source_path)
    manifest, manifest_identity = sealer._package_manifest(staging)
    staging_payload = {
        "schema": sealer.STAGING_CHECKPOINT_SCHEMA,
        "stagingRootName": staging.name,
        "sourceCheckpoint": _identity(source_path, source_path.parent),
        "sourceCheckpointPayloadSha256": contract.payload_sha256(source),
        "sourceAuthority": {
            "docwen": source["docwen"],
            "activeSemantics": source["activeSemantics"],
        },
        "packageManifest": manifest_identity,
        "packageManifestPayloadSha256": contract.payload_sha256(manifest),
        "stagingTree": sealer._payload_tree(staging),
    }
    staging_path = tmp_path / "external" / "staging.json"
    _write_json(staging_path, staging_payload)
    return source_path, source_sha, staging_path, contract.sha256_file(staging_path)


def test_v4_authority_and_schemas_freeze_docwen_identity() -> None:
    authority = json.loads((ROOT / contract.AUTHORITY_PATH).read_text(encoding="utf-8"))
    assert authority["activeSemantics"]["id"] == "docwen.markdown_semantics.v3"
    assert authority["docwenSpecBaseline"] == {
        "commit": "adebd0c93c5d2c16727d7db456c91504b54e099e",
        "tree": "1adab70aad02733adfd69930aaf39f473d773e43",
    }
    assert authority["docwenImplementationBaseline"]["status"] == "pending"
    assert "wenleafSpecBaseline" not in authority
    assert "wenleafImplementationBaseline" not in authority
    assert "assistantConsumer" not in authority
    assert {item["id"] for item in authority["exclusions"]} == {
        "docwen.markdown_semantics.v1",
        "docwen.markdown_semantics.v2",
        "docwen-v3-dirty-audit-head",
        "Attempt04",
    }
    for name in (contract.RECEIPT_SCHEMA_PATH, contract.EVIDENCE_SCHEMA_PATH):
        schema = json.loads((ROOT / name).read_text(encoding="utf-8"))
        Draft202012Validator.check_schema(schema)


def test_source_checkpoint_refuses_pending_docwen_implementation_and_ignored_inputs(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    docwen, final = _init_repo(tmp_path / "docwen")
    monkeypatch.setattr(
        contract,
        "load_authority",
        lambda _repo: (
            {
                "activeSemantics": {"id": contract.ACTIVE_SEMANTICS},
                "docwenSpecBaseline": docwen,
                "docwenImplementationBaseline": {"status": "pending", "commit": None, "tree": None},
            },
            {"relativePath": contract.AUTHORITY_PATH, "bytes": 1, "sha256": "0" * 64},
        ),
    )
    monkeypatch.setattr(contract, "final_spec_identity", lambda *_args: {})
    monkeypatch.setattr(contract, "oracle_manifest_identity", lambda *_args: {})
    with pytest.raises(contract.V4CandidateContractError, match="docwen_implementation_baseline_pending"):
        contract.capture_source_checkpoint(docwen_repo=tmp_path / "docwen")
    assert final == contract.clean_git_identity(tmp_path / "docwen", label="docwen")
    assert {".pyc", ".pyo"}.issubset(contract._EXECUTABLE_SUFFIXES)
    ignored = tmp_path / "docwen" / "ignored-resource.bin"
    ignored.write_bytes(b"not-an-executable-but-still-a-build-input")
    (tmp_path / "docwen" / ".gitignore").write_text("ignored-resource.bin\n", encoding="utf-8")
    assert contract.ignored_source_inputs(tmp_path / "docwen") == ["ignored-resource.bin"]


def test_source_checkpoint_rejects_docwen_ignored_inputs(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    identity = {"commit": "1" * 40, "tree": "2" * 40}
    authority = {
        "activeSemantics": {"id": contract.ACTIVE_SEMANTICS},
        "docwenSpecBaseline": identity,
        "docwenImplementationBaseline": {"status": "frozen", **identity},
    }
    docwen = tmp_path / "docwen"
    monkeypatch.setattr(contract, "load_authority", lambda _repo: (authority, {"sha256": "0" * 64}))
    monkeypatch.setattr(contract, "clean_git_identity", lambda *_args, **_kwargs: identity)
    monkeypatch.setattr(contract, "verify_baseline_ancestry", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(contract, "ignored_executable_inputs", lambda _repo: [])
    monkeypatch.setattr(contract, "ignored_source_inputs", lambda _repo: ["ignored.cache"])
    with pytest.raises(contract.V4CandidateContractError, match="docwen_ignored_source_input_present"):
        contract.capture_source_checkpoint(docwen_repo=docwen)


def test_seal_requires_external_hashes_and_emits_acyclic_noneligible_receipt(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    _, docwen = _init_repo(tmp_path / "docwen")
    candidate_id = f"docwen-0.9.0-v4-20260814T120000Z-{docwen['commit'][:12]}"
    staging = _staging(tmp_path, candidate_id)
    source = _source_payload(docwen, staging)
    monkeypatch.setattr(contract, "capture_source_checkpoint", lambda **_kwargs: source)
    source_path, source_sha, staging_path, staging_sha = _checkpoints(tmp_path, staging, source)
    monkeypatch.setattr(contract, "ignored_executable_inputs", lambda _repo: [])
    monkeypatch.setattr(contract, "ignored_source_inputs", lambda _repo: [])
    output = tmp_path / "output"
    output.mkdir()
    with pytest.raises(sealer.V4CandidateSealError, match="source_checkpoint_sha256_mismatch"):
        sealer.seal_candidate(
            docwen_repo=tmp_path / "docwen",
            staging_root=staging,
            source_checkpoint=source_path,
            source_checkpoint_sha256="0" * 64,
            staging_checkpoint=staging_path,
            staging_checkpoint_sha256=staging_sha,
            output_root=output,
            candidate_id=candidate_id,
            generated_at="2026-08-14T12:00:00Z",
        )
    result = sealer.seal_candidate(
        docwen_repo=tmp_path / "docwen",
        staging_root=staging,
        source_checkpoint=source_path,
        source_checkpoint_sha256=source_sha,
        staging_checkpoint=staging_path,
        staging_checkpoint_sha256=staging_sha,
        output_root=output,
        candidate_id=candidate_id,
        generated_at="2026-08-14T12:00:00Z",
    )
    final = Path(cast(str, result["candidateRoot"]))
    receipt = json.loads((final / "receipt.json").read_text(encoding="utf-8"))
    candidate = json.loads((final / "candidate.json").read_text(encoding="utf-8"))
    outer = json.loads((final / "manifest.json").read_text(encoding="utf-8"))
    assert receipt["consumerEligible"] is False
    assert receipt["state"] == "local_unpublished_host_validation_pending"
    assert "receipt" not in receipt["hashDag"] and "candidate" not in receipt["hashDag"]
    assert candidate["receipt"] == _identity(final / "receipt.json", final)
    assert outer["receipt"] == candidate["receipt"]
    assert outer["candidate"] == _identity(final / "candidate.json", final)
    assert receipt["evidence"]["sixIndependentInvalidIdRecords"] == contract.REQUIRED_INVALID_CASES
    assert candidate_evidence._capture_tree_stable(staging)["manifestSha256"]
    wrong_role = json.loads(json.dumps(receipt))
    wrong_role["evidence"]["records"][0]["identityPointers"][0]["role"] = "wire_schema"
    with pytest.raises(contract.V4CandidateContractError, match="json_schema_validation_failed"):
        contract.validate_json_schema(wrong_role, final / contract.RECEIPT_SCHEMA_PATH)
    not_run_with_record = json.loads(json.dumps(receipt))
    unexpected_host = not_run_with_record["evidence"]["records"][0]
    unexpected_host["caseId"] = "unexpected-word-host"
    unexpected_host["layer"] = "word_host"
    unexpected_host["identityPointers"][0]["role"] = "host_manifest"
    with pytest.raises(contract.V4CandidateContractError, match="json_schema_validation_failed"):
        contract.validate_json_schema(not_run_with_record, final / contract.RECEIPT_SCHEMA_PATH)


def test_all_host_layers_are_required_for_consumer_eligibility(tmp_path: Path) -> None:
    candidate_id = "docwen-0.9.0-v4-20260814T120000Z-0123456789ab"
    staging = _staging(tmp_path, candidate_id, host_status="passed")
    index = contract.read_json_object(staging / "_evidence/v4-evidence-index.json", label="index")
    records, _ = contract.validate_evidence_index(index, candidate_id=candidate_id)
    assert len(records) == 14
    assert contract.consumer_eligible(cast(Mapping[str, object], index["layerStatus"])) is True


def test_machine_consumer_contract_exposes_exact_closed_route_schemas() -> None:
    options = cast(Mapping[str, object], contract.machine_consumer_contract()["optionsSchemas"])
    docx = cast(Mapping[str, object], cast(Mapping[str, object], options["docx"])["schema"])
    fixed = cast(Mapping[str, object], cast(Mapping[str, object], options["pdfOfdXps"])["schema"])
    tiff = cast(Mapping[str, object], cast(Mapping[str, object], options["tiff"])["schema"])
    for schema in (docx, fixed, tiff):
        assert schema["type"] == "object"
        assert schema["required"] == []
        assert schema["additionalProperties"] is False
        properties = cast(Mapping[str, object], schema["properties"])
        assert properties["recognize_text"] == {"type": "boolean", "default": False}
        assert properties["preserve_resources"] == {"type": "boolean", "default": True}
    assert set(cast(Mapping[str, object], docx["properties"])) == {
        "recognize_text",
        "preserve_resources",
        "ocr_language",
        "image_mode",
        "ocr_placement",
        "image_link_style",
        "table_merge_strategy",
        "remove_numbering",
        "add_numbering",
        "numbering_scheme",
    }
    assert set(cast(Mapping[str, object], fixed["properties"])) == {
        "recognize_text",
        "preserve_resources",
        "ocr_language",
        "image_mode",
        "render_dpi",
    }
    assert set(cast(Mapping[str, object], tiff["properties"])) == {
        "recognize_text",
        "preserve_resources",
        "ocr_language",
    }


def test_external_checkpoint_writers_reject_source_or_staging_paths(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    docwen = tmp_path / "docwen"
    staging = tmp_path / "staging"
    for root in (docwen, staging):
        root.mkdir()
    monkeypatch.setattr(sealer, "_source_checkpoint_payload", lambda *_args: {"schema": "synthetic"})
    with pytest.raises(sealer.V4CandidateSealError, match="outside_source_repository"):
        sealer.write_source_checkpoint(
            docwen_repo=docwen,
            output=docwen / "checkpoint.json",
        )
    source = tmp_path / "source.json"
    source.write_text("{}\n", encoding="utf-8")
    with pytest.raises(sealer.V4CandidateSealError, match="outside_staging"):
        sealer.write_staging_checkpoint(
            docwen_repo=docwen,
            staging_root=staging,
            source_checkpoint=source,
            source_checkpoint_sha256=contract.sha256_file(source),
            output=staging / "checkpoint.json",
        )


def test_file_identity_rejects_linked_paths(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    target = tmp_path / "target.bin"
    target.write_bytes(b"target")
    linked = tmp_path / "linked.bin"
    try:
        linked.symlink_to(target)
    except OSError:
        monkeypatch.setattr(
            candidate_evidence,
            "_safe_regular_file",
            lambda *_args, **_kwargs: (_ for _ in ()).throw(candidate_evidence.EvidenceError("linked_path_rejected")),
        )
    with pytest.raises(contract.V4CandidateContractError, match="linked_path_rejected"):
        contract.file_identity(linked, relative_to=tmp_path)


def test_external_checkpoint_rejects_symlink_before_resolution(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    target = tmp_path / "target.json"
    _write_json(target, {"schema": "synthetic"})
    target_sha = contract.sha256_file(target)
    linked = tmp_path / "linked.json"
    try:
        linked.symlink_to(target)
    except OSError:
        monkeypatch.setattr(
            candidate_evidence,
            "_safe_regular_file",
            lambda *_args, **_kwargs: (_ for _ in ()).throw(candidate_evidence.EvidenceError("linked_path_rejected")),
        )
    with pytest.raises(sealer.V4CandidateSealError, match="linked_path_rejected"):
        sealer._read_external_checkpoint(linked, expected_sha256=target_sha, label="checkpoint")


def test_atomic_publish_quarantines_final_on_post_rename_proof_failure(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    publishing = tmp_path / ".publishing-candidate"
    publishing.mkdir()
    (publishing / "payload.bin").write_bytes(b"stable")
    expected = candidate_evidence._capture_tree_stable(publishing)
    final = tmp_path / "candidate"
    original = candidate_evidence._capture_tree_stable
    calls = 0

    def fail_final(root: Path, *args: Any, **kwargs: Any) -> dict[str, object]:
        nonlocal calls
        calls += 1
        if root == final:
            raise candidate_evidence.EvidenceError("synthetic_post_publish_failure")
        return original(root, *args, **kwargs)

    monkeypatch.setattr(candidate_evidence, "_capture_tree_stable", fail_final)
    with pytest.raises(sealer.V4CandidateSealError, match="candidate_quarantined_after_failed_proof"):
        sealer._atomic_publish(publishing_root=publishing, final_root=final, expected_tree=expected)
    assert calls >= 2
    assert not final.exists()
    assert len(list(tmp_path.glob(".rejected-candidate-*"))) == 1


def test_atomic_publish_reports_quarantine_failure_and_leaves_nonreusable_canonical(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    publishing = tmp_path / ".publishing-candidate"
    publishing.mkdir()
    (publishing / "payload.bin").write_bytes(b"stable")
    expected = candidate_evidence._capture_tree_stable(publishing)
    final = tmp_path / "candidate"
    original_capture = candidate_evidence._capture_tree_stable
    original_rename = Path.rename

    def fail_final(root: Path, *args: Any, **kwargs: Any) -> dict[str, object]:
        if root == final:
            raise candidate_evidence.EvidenceError("synthetic_post_publish_failure")
        return original_capture(root, *args, **kwargs)

    def refuse_quarantine(path: Path, target: Path) -> Path:
        if target.name.startswith(".rejected-"):
            raise OSError("synthetic_quarantine_failure")
        return original_rename(path, target)

    monkeypatch.setattr(candidate_evidence, "_capture_tree_stable", fail_final)
    monkeypatch.setattr(Path, "rename", refuse_quarantine)
    with pytest.raises(sealer.V4CandidateSealError, match="post_publish_proof_failed_and_quarantine_failed"):
        sealer._atomic_publish(publishing_root=publishing, final_root=final, expected_tree=expected)
    assert final.exists()
    with pytest.raises(sealer.V4CandidateSealError, match="final_candidate_path_exists"):
        sealer._atomic_publish(publishing_root=publishing, final_root=final, expected_tree=expected)


def test_candidate_tool_files_remain_under_governance_threshold() -> None:
    for relative in (
        "scripts/release/v4_candidate_contract.py",
        "scripts/release/seal_v4_candidate.py",
        "tests/test_repo/test_v4_candidate_receipt.py",
    ):
        assert len((ROOT / relative).read_text(encoding="utf-8").splitlines()) < 700
