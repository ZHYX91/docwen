from __future__ import annotations

import copy
import hashlib
import json
import os
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pytest
from scripts.release import build_v4_candidate_staging as staging_builder
from scripts.release import build_v4_package_input as producer
from scripts.release import v4_candidate_contract as contract
from tests.test_repo import v4_package_input_test_support as support
from tests.test_repo import v4_package_input_transcript_test_support as transcript_support

pytestmark = pytest.mark.contract

DOCWEN_COMMIT = "a" * 40
DOCWEN_TREE = "b" * 40
CANDIDATE_ID = f"docwen-0.9.0-v4-20260814T120000Z-{DOCWEN_COMMIT[:12]}"
SOURCE = "# 2.3 Authored heading ^heading-one\n\nTable: Caption ^table-one\n\nSee @[[#^table-one]].\n"


@dataclass(frozen=True)
class SyntheticFixture:
    docwen: Path
    checkpoint: dict[str, Any]
    checkpoint_path: Path
    checkpoint_sha256: str
    output: Path
    work: Path


def _json_bytes(value: object) -> bytes:
    return contract.json_bytes(value)


def _write_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(_json_bytes(value))


def _identity(path: Path, root: Path) -> dict[str, object]:
    return contract.file_identity(path, relative_to=root)


def _case(case_id: str, source: str, diagnostics: list[dict[str, object]]) -> dict[str, object]:
    return {
        "case_id": case_id,
        "category": "negative" if diagnostics else "positive",
        "input_id": f"{case_id}.md",
        "source": source,
        "expected": {
            "target_kind_ids": [],
            "anchor_kind_ids": [],
            "link_kinds": [],
            "reference_statuses": [],
            "citation_keys": [],
            "diagnostics": diagnostics,
            "source_unchanged": True,
        },
    }


def _source_cases(docwen: Path) -> None:
    cases = {
        "invalid-id-dot": ("Paragraph ^bad.id\n", "docwen.markdown.anchor.invalid_id", 10, 17),
        "invalid-id-slash": ("Paragraph ^bad/id\n", "docwen.markdown.anchor.invalid_id", 10, 17),
        "invalid-id-underscore": ("Paragraph ^bad_id\n", "docwen.markdown.anchor.invalid_id", 10, 17),
        "invalid-id-too-long": ("Paragraph ^" + "a" * 129 + "\n", "docwen.markdown.anchor.invalid_id", 10, 140),
        "invalid-duplicate": ("A ^same\n\nB ^same\n", "docwen.markdown.anchor.duplicate_id", 12, 17),
        "invalid-kind-mismatch": (
            "Figure: Caption ^same\n\n@[[#^same]]\n",
            "docwen.markdown.cross_reference.kind_mismatch",
            23,
            33,
        ),
    }
    corpus = docwen / producer.SOURCE_ORACLE_ROOT / "corpus"
    for case_id, (source, code, start, end) in cases.items():
        _write_json(
            corpus / f"{case_id}.case.json",
            _case(case_id, source, [{"severity": "error", "code": code, "range": {"start": start, "end": end}}]),
        )
    _write_json(corpus / "positive-exact-two.case.json", _case("positive-exact-two", SOURCE, []))
    rows: list[dict[str, object]] = []
    for path in sorted(corpus.glob("*.json")):
        identity = _identity(path, docwen / producer.SOURCE_ORACLE_ROOT)
        rows.append(
            {
                "path": identity["relativePath"],
                "bytes": identity["bytes"],
                "sha256": identity["sha256"],
            }
        )
    _write_json(
        docwen / producer.SOURCE_ORACLE_MANIFEST,
        {
            "schema": "docwen.markdown_semantics_manifest.v1",
            "oracle": {"semantics": "docwen.markdown_semantics.v3"},
            "evidence_layer": "source_oracle",
            "files": rows,
        },
    )


def _harness_files(docwen: Path) -> None:
    root = docwen / Path(producer.HARNESS_MANIFEST_RELATIVE).parent
    source_path = root / "source/exact-two.md"
    source_path.parent.mkdir(parents=True, exist_ok=True)
    source_path.write_bytes(SOURCE.encode())
    cases: list[dict[str, object]] = []
    for case_id in producer.REQUIRED_HARNESS_CASE_IDS:
        neutral, plan, expected = support.case_semantics(case_id, SOURCE)
        neutral_path = root / f"expected-docwen/{case_id}.resolved.json"
        plan_path = root / f"expected-docwen/{case_id}.plan.json"
        _write_json(neutral_path, neutral)
        _write_json(plan_path, plan)
        coverage = [case_id]
        if case_id == "rich-semantics-composite":
            coverage = sorted(
                [
                    "five-kind-enabled",
                    "semantic-cross-reference",
                    "citation",
                    "embedded-raster",
                    "semantic-bibliography",
                    "ordinary-wikilink-preserved",
                ]
            )
        cases.append(
            {
                "caseId": case_id,
                "sourceOracleCaseId": "positive-exact-two",
                "coverage": coverage,
                "source": _identity(source_path, docwen),
                "neutralDocument": _identity(neutral_path, docwen),
                "numberingExportPlan": _identity(plan_path, docwen),
                "expectedOoxml": expected,
                "expectedRoundtrip": _identity(source_path, docwen),
            }
        )
    _write_json(
        docwen / producer.HARNESS_MANIFEST_RELATIVE,
        {
            "schema": producer.HARNESS_SCHEMA,
            "harness": {"id": producer.HARNESS_ID, "version": producer.HARNESS_VERSION},
            "requiredCaseIds": list(producer.REQUIRED_HARNESS_CASE_IDS),
            "pendingCaseIds": [],
            "cases": cases,
        },
    )


def _fixture(tmp_path: Path) -> SyntheticFixture:
    repositories = tmp_path / "repositories"
    docwen = repositories / "docwen"
    docwen.mkdir(parents=True)
    _source_cases(docwen)
    _harness_files(docwen)
    wire = docwen / contract.WIRE_SCHEMA_PATH
    _write_json(wire, {"$id": "urn:docwen:schema:machine-diagnostic-evidence:v1"})
    oracle_identity = _identity(docwen / producer.SOURCE_ORACLE_MANIFEST, docwen)
    wire_identity = _identity(wire, docwen)
    checkpoint: dict[str, Any] = {
        "schema": "docwen.v4_candidate_source_checkpoint.v1",
        "authorityBaseline": {"relativePath": "authority.json", "bytes": 1, "sha256": "0" * 64},
        "activeSemantics": {"id": "docwen.markdown_semantics.v3"},
        "finalSpecIdentity": {"files": [wire_identity], "sha256": "1" * 64},
        "oracleManifest": oracle_identity,
        "docwen": {
            "specBaseline": {"commit": "1" * 40, "tree": "2" * 40},
            "implementationBaseline": {"commit": "3" * 40, "tree": "4" * 40},
            "final": {"commit": DOCWEN_COMMIT, "tree": DOCWEN_TREE},
        },
    }
    checkpoint_path = tmp_path / "external/source-checkpoint.json"
    _write_json(checkpoint_path, checkpoint)
    return SyntheticFixture(
        docwen=docwen,
        checkpoint=checkpoint,
        checkpoint_path=checkpoint_path,
        checkpoint_sha256=contract.sha256_file(checkpoint_path),
        output=tmp_path / "candidate-input",
        work=tmp_path / "disposable-work",
    )


def _checkpoint_loader(fixture: SyntheticFixture) -> producer.CheckpointLoader:
    def load(path: Path, digest: str, docwen: Path) -> tuple[dict[str, Any], dict[str, object]]:
        assert path == fixture.checkpoint_path
        assert digest == fixture.checkpoint_sha256
        assert os.path.samefile(docwen, fixture.docwen)
        return copy.deepcopy(fixture.checkpoint), _identity(path, path.parent)

    return load


def _clone(source: Path, destination: Path, commit: str, tree: str, label: str) -> None:
    assert label == "docwen"
    assert len(commit) == len(tree) == 40
    shutil.copytree(source, destination)


def _package_builder(
    different_cli: bool = False,
) -> producer.PackageBuilder:
    def build(clone: Path, python: Path, uv: Path) -> producer.BuildOutput:
        assert python.name == "python.exe"
        assert uv.name == "uv.exe"
        gui = clone / "synthetic-dist" / producer.PACKAGE_NAMES[0]
        cli = clone / "synthetic-dist" / producer.PACKAGE_NAMES[1]
        gui.mkdir(parents=True)
        cli.mkdir()
        (gui / "DocWen.exe").write_bytes(b"synthetic-gui")
        (gui / "DocWenCLI.exe").write_bytes(b"synthetic-cli")
        (cli / "DocWenCLI.exe").write_bytes(b"different-cli" if different_cli else b"synthetic-cli")
        return producer.BuildOutput(
            gui=gui,
            cli=cli,
            metadata={"command": ["synthetic-exact-build"], "toolchain": {"synthetic": True}},
        )

    return build


def _docx(case: Any) -> bytes:
    return support.build_docx(case.neutral_envelope, case.plan_envelope)


def _validation_terminal(docwen_clone: Path) -> dict[str, object]:
    fixture = json.loads(
        (docwen_clone / producer.SOURCE_ORACLE_ROOT / "corpus/invalid-id-dot.case.json").read_text(encoding="utf-8")
    )
    source = fixture["source"].encode()
    diagnostics = [
        {
            "severity": item["severity"],
            "code": item["code"],
            "message": "Synthetic source diagnostic.",
            "evidence_schema": "docwen.machine.diagnostic_evidence.v1",
            "source": {
                "input_id": "invalid-id-dot.md",
                "sha256": hashlib.sha256(source).hexdigest(),
                "encoding": "utf-8",
                "coordinate_system": "unicode_code_point",
                "offset_base": 0,
                "range_end": "exclusive",
            },
            "range": item["range"],
            "related_ranges": [],
            "fixes": [],
        }
        for item in fixture["expected"]["diagnostics"]
    ]
    task_id = "task.synthetic.validation"
    return {
        "jsonrpc": "2.0",
        "method": "task/completed",
        "params": {
            "task_id": task_id,
            "sequence": 1,
            "bundle": {
                "schema": "docwen.artifact_bundle.v2",
                "layout_schema": "docwen.artifact_layout.v1",
                "bundle_id": "bundle.synthetic.validation",
                "task_id": task_id,
                "producer": {
                    "name": "DocWen",
                    "product_version": "0.9.0",
                    "machine_protocol": "docwen.machine.v1",
                },
                "artifacts": [
                    {
                        "artifact_id": "document.validation",
                        "kind": "document",
                        "locator": "validation/report.json",
                        "logical_path": "validation/report.json",
                        "suggested_name": "report.json",
                        "media_type": "application/json",
                        "size_bytes": 1,
                        "sha256": "9" * 64,
                    }
                ],
                "entries": [{"artifact_id": "document.validation", "role": "primary", "ordinal": 0, "preferred": True}],
                "relations": [],
            },
            "diagnostics": diagnostics,
            "metrics": {"duration_ms": 1, "input_bytes": len(source), "output_bytes": 1},
        },
    }


def _harness_runner(*, roundtrip: bytes | None = None) -> producer.HarnessRunner:
    def run(
        executable: Path,
        docwen_clone: Path,
        run_root: Path,
        harness: producer.HarnessInput,
    ) -> producer.HarnessOutput:
        assert executable.name == "DocWenCLI.exe"
        assert not run_root.exists()
        case_outputs: list[producer.HarnessCaseOutput] = []
        for case in harness.cases:
            docx = _docx(case)
            case_outputs.append(
                producer.HarnessCaseOutput(
                    case_id=case.case_id,
                    docx=docx,
                    roundtrip=case.expected_roundtrip if roundtrip is None else roundtrip,
                    inspection=producer._inspect_docx(
                        docx,
                        case.expected_ooxml,
                        case.neutral_envelope,
                        case.plan_envelope,
                    ),
                )
            )
        outputs = tuple(case_outputs)
        terminal = _validation_terminal(docwen_clone)
        transcript, stdout, request_digest = transcript_support.synthetic_transcript(harness, outputs, terminal)
        return producer.HarnessOutput(
            validation_terminal=terminal,
            stdout=stdout,
            stderr=b"",
            transcript=transcript,
            cases=outputs,
            request_digest=request_digest,
        )

    return run


def _arguments(fixture: SyntheticFixture, **overrides: object) -> dict[str, object]:
    values: dict[str, object] = {
        "docwen_repo": fixture.docwen,
        "source_checkpoint": fixture.checkpoint_path,
        "source_checkpoint_sha256": fixture.checkpoint_sha256,
        "candidate_id": CANDIDATE_ID,
        "python": fixture.checkpoint_path.parent / "python.exe",
        "uv": fixture.checkpoint_path.parent / "uv.exe",
        "output_root": fixture.output,
        "work_root": fixture.work,
        "checkpoint_loader": _checkpoint_loader(fixture),
        "clone_factory": _clone,
        "clone_verifier": lambda *_arguments: None,
        "package_builder": _package_builder(),
        "harness_runner": _harness_runner(),
        "version_reader": lambda _path: "DocWen 0.9.0 (CLI protocol 3)",
    }
    values.update(overrides)
    return values


def test_builds_closed_exact_two_inputs_and_honest_not_run_hosts(tmp_path: Path) -> None:
    fixture = _fixture(tmp_path)
    result = producer.build_package_input(**_arguments(fixture))  # type: ignore[arg-type]

    assert os.path.samefile(Path(result["outputRoot"]), fixture.output)
    plan = json.loads((fixture.output / "evidence-plan.json").read_text(encoding="utf-8"))
    statuses, records, manifests = staging_builder._parse_plan(plan, candidate_id=CANDIDATE_ID)
    assert all(statuses[layer] == "not_run" for layer in contract.REQUIRED_LAYERS[6:])
    assert all(not any(item["layer"] == layer for item in records) for layer in contract.REQUIRED_LAYERS[6:])
    assert all(manifests[layer] is None for layer in contract.REQUIRED_LAYERS[6:])
    assert all(
        item["layer"] == "source_oracle" or item["caseId"].startswith(producer.HARNESS_CASE_PREFIX) for item in records
    )
    assert len([item for item in records if item["layer"] == "packaged"]) == 1
    assert len([item for item in records if item["layer"] == "roundtrip"]) == len(producer.REQUIRED_HARNESS_CASE_IDS)
    assert len([item for item in records if item["layer"] == "headless_ooxml"]) == len(
        producer.REQUIRED_HARNESS_CASE_IDS
    )
    result_payload = json.loads((fixture.output / "producer-result.json").read_text(encoding="utf-8"))
    assert result_payload["harness"]["id"] == producer.HARNESS_ID
    assert result_payload["harness"]["version"] == producer.HARNESS_VERSION
    assert result_payload["harness"]["requiredCaseIds"] == list(producer.REQUIRED_HARNESS_CASE_IDS)
    assert result_payload["harness"]["executedCaseIds"] == list(producer.REQUIRED_HARNESS_CASE_IDS)
    assert result_payload["harness"]["pendingCaseIds"] == []
    assert result_payload["harness"]["forwardInputRoles"] == ["neutral_document", "numbering_export_plan"]
    assert result_payload["package"]["cliBytesIdentical"] is True
    assert result_payload["package"]["standaloneCli"]["bytes"] == len(b"synthetic-cli")
    assert result_payload["package"]["standaloneCli"]["version"] == "DocWen 0.9.0 (CLI protocol 3)"
    assert result_payload["evidence"]["hostStatus"] == {
        "word_host": "not_run",
        "wps_host": "not_run",
        "libreoffice_host": "not_run",
    }


def test_exact_two_request_rejects_cardinality_role_media_and_hash_mutations(tmp_path: Path) -> None:
    neutral = tmp_path / "neutral.json"
    plan = tmp_path / "plan.json"
    neutral.write_bytes(b"{}")
    plan.write_bytes(b"{}")
    canonical = producer._exact_two_inputs(neutral, plan)
    producer._validate_exact_two_request(canonical, {})

    mutations: list[tuple[list[dict[str, object]], object]] = []
    mutations.append((canonical[:1], {}))
    extra = copy.deepcopy(canonical)
    extra.append(copy.deepcopy(canonical[0]))
    mutations.append((extra, {}))
    old_roles = copy.deepcopy(canonical)
    old_roles[0]["role"] = "source"
    old_roles[1]["role"] = "bibliography"
    mutations.append((old_roles, {}))
    wrong_media = copy.deepcopy(canonical)
    wrong_media[1]["media_type"] = "application/json"
    mutations.append((wrong_media, {}))
    wrong_hash = copy.deepcopy(canonical)
    wrong_hash[0]["sha256"] = "0" * 64
    mutations.append((wrong_hash, {}))
    extra_key = copy.deepcopy(canonical)
    extra_key[0]["legacy"] = True
    mutations.append((extra_key, {}))
    mutations.append((copy.deepcopy(canonical), {"add_numbering": True}))
    for inputs, options in mutations:
        with pytest.raises(producer.V4PackageInputError):
            producer._validate_exact_two_request(inputs, options)


def test_harness_manifest_rejects_attempt04_and_c6_tokens(tmp_path: Path) -> None:
    fixture = _fixture(tmp_path)
    manifest_path = fixture.docwen / producer.HARNESS_MANIFEST_RELATIVE
    for legacy in ("Attempt04", "C6-V", "c6v"):
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        manifest["cases"][0]["source"]["relativePath"] = f"tests/fixtures/document-semantics-v4/source/{legacy}.md"
        _write_json(manifest_path, manifest)
        with pytest.raises(producer.V4PackageInputError, match="legacy_harness_marker_rejected"):
            producer._load_harness(fixture.docwen)
        _harness_files(fixture.docwen)

    for legacy in ("source+bibliography", "source-bibliography", "machine.typed-markdown-png"):
        with pytest.raises(producer.V4PackageInputError, match="legacy_harness_marker_rejected"):
            producer._reject_legacy_harness({"legacy": legacy})


def test_stale_harness_pointer_fails_closed(tmp_path: Path) -> None:
    fixture = _fixture(tmp_path)
    manifest_path = fixture.docwen / producer.HARNESS_MANIFEST_RELATIVE
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    manifest["cases"][0]["numberingExportPlan"]["sha256"] = "0" * 64
    _write_json(manifest_path, manifest)
    with pytest.raises(producer.V4PackageInputError, match="pointer_identity_mismatch"):
        producer._load_harness(fixture.docwen)


def test_harness_matrix_pending_order_coverage_and_roundtrip_mutations_fail_closed(tmp_path: Path) -> None:
    fixture = _fixture(tmp_path)
    manifest_path = fixture.docwen / producer.HARNESS_MANIFEST_RELATIVE
    original = json.loads(manifest_path.read_text(encoding="utf-8"))

    pending = copy.deepcopy(original)
    pending_id = producer.REQUIRED_HARNESS_CASE_IDS[-1]
    pending["pendingCaseIds"] = [pending_id]
    pending["cases"] = pending["cases"][:-1]
    _write_json(manifest_path, pending)
    with pytest.raises(producer.V4PackageInputError, match="harness_cases_pending"):
        producer._load_harness(fixture.docwen)

    duplicated = copy.deepcopy(original)
    duplicated["cases"].append(copy.deepcopy(duplicated["cases"][0]))
    _write_json(manifest_path, duplicated)
    with pytest.raises(producer.V4PackageInputError, match="case_set_or_order_invalid"):
        producer._load_harness(fixture.docwen)

    incomplete = copy.deepcopy(original)
    rich = next(item for item in incomplete["cases"] if item["caseId"] == "rich-semantics-composite")
    rich["coverage"].remove("citation")
    _write_json(manifest_path, incomplete)
    with pytest.raises(producer.V4PackageInputError, match="rich_coverage_incomplete"):
        producer._load_harness(fixture.docwen)

    stale_roundtrip = copy.deepcopy(original)
    stale_roundtrip["cases"][0]["expectedRoundtrip"]["sha256"] = "0" * 64
    _write_json(manifest_path, stale_roundtrip)
    with pytest.raises(producer.V4PackageInputError, match="pointer_identity_mismatch"):
        producer._load_harness(fixture.docwen)


def test_binary_difference_and_version_drift_are_rejected(tmp_path: Path) -> None:
    different = _fixture(tmp_path / "different")
    with pytest.raises(producer.V4PackageInputError, match="cli_bytes_differ"):
        producer.build_package_input(
            **_arguments(different, package_builder=_package_builder(different_cli=True))  # type: ignore[arg-type]
        )
    assert not different.output.exists()

    drift = _fixture(tmp_path / "drift")
    with pytest.raises(producer.V4PackageInputError, match="cli_versions_differ"):
        producer.build_package_input(
            **_arguments(  # type: ignore[arg-type]
                drift,
                version_reader=lambda path: (
                    "DocWen 0.9.1 (CLI protocol 3)"
                    if producer.PACKAGE_NAMES[0] in path.parts
                    else "DocWen 0.9.0 (CLI protocol 3)"
                ),
            )
        )
    assert not drift.output.exists()


def test_roundtrip_mismatch_is_quarantined_without_publishing(tmp_path: Path) -> None:
    fixture = _fixture(tmp_path)
    with pytest.raises(producer.V4PackageInputError, match="roundtrip_not_byte_exact"):
        producer.build_package_input(
            **_arguments(fixture, harness_runner=_harness_runner(roundtrip=b"changed\n"))  # type: ignore[arg-type]
        )
    assert not fixture.output.exists()
    assert len(list(fixture.output.parent.glob(f".{fixture.output.name}.rejected-*"))) == 1


def test_missing_executed_case_is_quarantined(tmp_path: Path) -> None:
    fixture = _fixture(tmp_path)
    complete_runner = _harness_runner()

    def missing_case(
        executable: Path,
        docwen_clone: Path,
        run_root: Path,
        harness: producer.HarnessInput,
    ) -> producer.HarnessOutput:
        output = complete_runner(executable, docwen_clone, run_root, harness)
        return producer.HarnessOutput(
            validation_terminal=output.validation_terminal,
            stdout=output.stdout,
            stderr=output.stderr,
            transcript=output.transcript,
            cases=output.cases[:-1],
            request_digest=output.request_digest,
        )

    with pytest.raises(producer.V4PackageInputError, match="transcript_output_case_set_invalid"):
        producer.build_package_input(  # type: ignore[arg-type]
            **_arguments(fixture, harness_runner=missing_case)
        )
    assert not fixture.output.exists()


def test_post_publish_stability_failure_quarantines_output(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    fixture = _fixture(tmp_path)
    capture = producer.evidence_contract.capture_tree_stable

    def drift_after_publish(path: Path) -> dict[str, object]:
        identity = capture(path)
        if path.exists() and fixture.output.exists() and os.path.samefile(path, fixture.output):
            return {**identity, "manifestSha256": "f" * 64}
        return identity

    monkeypatch.setattr(producer.evidence_contract, "capture_tree_stable", drift_after_publish)
    with pytest.raises(producer.V4PackageInputError, match="producer_output_changed_after_publish"):
        producer.build_package_input(**_arguments(fixture))  # type: ignore[arg-type]
    assert not fixture.output.exists()
    assert len(list(fixture.output.parent.glob(f".{fixture.output.name}.rejected-*"))) == 1


def test_headless_inspection_requires_numbering_and_exact_seq_ref_bookmarks(tmp_path: Path) -> None:
    fixture = _fixture(tmp_path)
    harness = producer._load_harness(fixture.docwen)
    case = next(item for item in harness.cases if item.case_id == "rich-semantics-composite")
    docx = _docx(case)
    assert producer._inspect_docx(
        docx,
        case.expected_ooxml,
        case.neutral_envelope,
        case.plan_envelope,
    ) == {
        "bookmarkCount": 6,
        "seqFieldCount": 4,
        "refFieldCount": 2,
        "violations": [],
    }
    with pytest.raises(producer.V4PackageInputError, match="headless_ooxml_inspection_failed"):
        producer._inspect_docx(
            docx,
            {**case.expected_ooxml, "refFieldCount": 3},
            case.neutral_envelope,
            case.plan_envelope,
        )


def test_missing_fixed_harness_path_never_searches_for_an_alternative(tmp_path: Path) -> None:
    fixture = _fixture(tmp_path)
    expected = fixture.docwen / producer.HARNESS_MANIFEST_RELATIVE
    alternate = expected.parent / "Attempt04.json"
    alternate.parent.mkdir(parents=True, exist_ok=True)
    alternate.write_bytes(expected.read_bytes())
    expected.unlink()
    with pytest.raises(producer.V4PackageInputError):
        producer._load_harness(fixture.docwen)
