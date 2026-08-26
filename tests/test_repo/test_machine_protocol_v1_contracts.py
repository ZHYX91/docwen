from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path

import pytest
from jsonschema.exceptions import ValidationError

pytestmark = pytest.mark.contract

EXPECTED_SEMANTIC_CODES = {
    "bundle": {
        "dangling_entry",
        "dangling_relation",
        "duplicate_artifact_id",
        "duplicate_artifact_locator",
        "duplicate_entry",
        "duplicate_entry_ordinal",
        "duplicate_relation_ordinal",
        "duplicate_page_index",
        "incompatible_entry_role",
        "incompatible_relation",
        "invalid_suggested_name",
        "incomplete_page_sequence",
        "invalid_page_range",
        "missing_relation_ordinal",
        "missing_page_fragment_semantics",
        "multiple_structural_owners",
        "orphan_artifact",
        "owned_entry",
        "page_count_mismatch",
        "page_ordinal_mismatch",
        "page_source_mismatch",
        "preferred_entry_count",
        "relation_cycle",
        "self_relation",
        "resource_page_mismatch",
        "unexpected_page_semantics",
    },
    "framing": {
        "frame_length_mismatch",
        "frame_too_large",
        "invalid_frame_header",
        "invalid_frame_payload",
    },
    "trace": {
        "bundle_task_mismatch",
        "dangling_diagnostic_artifact",
        "duplicate_input_logical_path",
        "duplicate_request_id",
        "duplicate_task_id",
        "input_slot_cardinality_mismatch",
        "input_slot_kind_mismatch",
        "input_slot_media_type_mismatch",
        "missing_terminal",
        "nonmonotonic_sequence",
        "notification_after_terminal",
        "notification_before_acceptance",
        "unknown_cancel_task",
        "unmatched_request",
        "unmatched_response",
        "undeclared_input_role",
    },
}


def _load_validator_module():
    project_root = Path(__file__).resolve().parents[2]
    script_path = project_root / "tools" / "validate_contracts.py"
    spec = importlib.util.spec_from_file_location("docwen_validate_contracts", script_path)
    assert spec is not None
    assert spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def test_machine_protocol_v1_contract_set_is_conformant() -> None:
    validator = _load_validator_module()
    contracts_root = Path(__file__).resolve().parents[2] / "contracts"

    summary = validator.validate_contract_set(contracts_root)

    assert summary.schemas == 9
    assert summary.valid_fixtures == 18
    assert summary.invalid_fixtures == 62
    manifest = validator_json(contracts_root / "conformance-manifest.json")
    assert {(record["name"], record["id"], record["path"]) for record in manifest["schemas"]} == {
        (
            "artifact_bundle",
            "urn:docwen:schema:artifact-bundle:v2",
            "schemas/docwen.artifact_bundle.v2.schema.json",
        ),
        (
            "machine",
            "urn:docwen:schema:machine-protocol:v1",
            "schemas/docwen.machine.v1.schema.json",
        ),
        (
            "machine_diagnostic_evidence",
            "urn:docwen:schema:machine-diagnostic-evidence:v1",
            "schemas/docwen.machine.diagnostic_evidence.v1.schema.json",
        ),
        (
            "numbering_export_plan",
            "urn:docwen:schema:numbering-export-plan:v1",
            "schemas/docwen.numbering_export_plan.v1.schema.json",
        ),
        (
            "proofread_report",
            "urn:docwen:schema:proofread-report:v2",
            "schemas/docwen.proofread_report.v2.schema.json",
        ),
        (
            "candidate_evidence_index",
            "urn:docwen:schema:candidate-evidence-index:v4",
            "schemas/docwen.candidate_evidence_index.v4.schema.json",
        ),
        (
            "candidate_receipt",
            "urn:docwen:schema:candidate-receipt:v4",
            "schemas/docwen.candidate_receipt.v4.schema.json",
        ),
        (
            "resolved_document",
            "urn:docwen:schema:resolved-document:v1",
            "schemas/docwen.resolved_document.v1.schema.json",
        ),
        (
            "semantic_bibliography",
            "urn:docwen:schema:semantic-bibliography:v1",
            "schemas/docwen.semantic_bibliography.v1.schema.json",
        ),
    }


def test_resolved_numbering_capability_is_exact_two_and_rejects_legacy_shape() -> None:
    validator = _load_validator_module()
    contracts_root = Path(__file__).resolve().parents[2] / "contracts"
    trace = validator_json(contracts_root / "fixtures" / "valid" / "machine.resolved-numbering.trace.json")
    capability = trace[1]["result"]["capabilities"][0]

    assert capability["input_shape"]["slots"] == [
        {
            "role": "neutral_document",
            "kind": "document",
            "media_types": ["application/vnd.docwen.resolved-document+json"],
            "min_items": 1,
            "max_items": 1,
        },
        {
            "role": "numbering_export_plan",
            "kind": "resource",
            "media_types": ["application/vnd.docwen.numbering-export-plan+json"],
            "min_items": 1,
            "max_items": 1,
        },
    ]
    assert [item["role"] for item in trace[2]["params"]["inputs"]] == [
        "neutral_document",
        "numbering_export_plan",
    ]
    assert {
        "remove_numbering",
        "add_numbering",
        "numbering_scheme",
        "heading_numbering_render_mode",
    }.isdisjoint(capability["options_schema"]["properties"])

    source_shape = json.loads(json.dumps(trace))
    source_shape[1]["result"]["capabilities"][0]["input_shape"]["slots"][0]["role"] = "source"
    with pytest.raises(validator.SemanticContractError) as rejected_shape:
        validator.validate_trace(source_shape, requires_terminal=False)
    assert rejected_shape.value.code == "invalid_resolved_numbering_slots"

    legacy_options = json.loads(json.dumps(trace))
    legacy_options[1]["result"]["capabilities"][0]["options_schema"]["properties"]["add_numbering"] = {
        "type": "boolean",
        "default": False,
    }
    with pytest.raises(validator.SemanticContractError) as rejected_options:
        validator.validate_trace(legacy_options, requires_terminal=False)
    assert rejected_options.value.code == "legacy_numbering_option_exposed"


def test_semantic_fixture_inventory_covers_each_normative_error_once() -> None:
    contracts_root = Path(__file__).resolve().parents[2] / "contracts"
    manifest = validator_json(contracts_root / "conformance-manifest.json")

    for document_type, expected_codes in EXPECTED_SEMANTIC_CODES.items():
        actual_codes = [
            record["error_code"]
            for record in manifest["fixtures"]
            if record["document_type"] == document_type and record["expect"] == "invalid_semantic"
        ]
        assert len(actual_codes) == len(set(actual_codes))
        assert set(actual_codes) == expected_codes


def test_framing_oracle_mismatch_is_a_harness_error_not_a_negative_fixture() -> None:
    validator = _load_validator_module()
    contracts_root = Path(__file__).resolve().parents[2] / "contracts"
    payload = validator_json(contracts_root / "fixtures" / "valid" / "machine.framing.json")
    payload["expected_messages"] = []

    with pytest.raises(validator.SemanticContractError) as rejected:
        validator.validate_framing_fixture(payload)

    assert rejected.value.code == "framing_oracle_mismatch"
    manifest = validator_json(contracts_root / "conformance-manifest.json")
    assert all(record.get("error_code") != "framing_oracle_mismatch" for record in manifest["fixtures"])


def test_route_semantics_are_consumer_neutral() -> None:
    contracts_root = Path(__file__).resolve().parents[2] / "contracts"
    gongwen = validator_json(contracts_root / "fixtures" / "valid" / "artifact-bundle.gongwen.json")
    ocr = validator_json(contracts_root / "fixtures" / "valid" / "artifact-bundle.ocr.json")
    worksheets = validator_json(contracts_root / "fixtures" / "valid" / "artifact-bundle.worksheets.json")

    assert {relation["type"] for relation in gongwen["relations"]} == {"attachment_of", "resource_of"}
    assert {artifact["kind"] for artifact in ocr["artifacts"]} == {"document", "fragment", "resource"}
    assert {relation["role"] for relation in ocr["relations"] if relation["type"] == "fragment_of"} == {"ocr_page"}
    page_relations = [relation for relation in ocr["relations"] if relation["role"] == "ocr_page"]
    assert [relation["page_fragment"]["page_index"] for relation in page_relations] == [1, 2, 3, 4]
    assert [relation["page_fragment"]["ocr_status"] for relation in page_relations] == [
        "success",
        "no_text",
        "recognition_failed",
        "success",
    ]
    assert len([artifact for artifact in ocr["artifacts"] if artifact["kind"] == "fragment"]) == 4
    assert len([artifact for artifact in ocr["artifacts"] if artifact["kind"] == "resource"]) == 5
    assert [entry["role"] for entry in worksheets["entries"]] == ["worksheet", "worksheet"]

    serialized = str([gongwen, ocr, worksheets]).lower()
    for forbidden in ("wenleaf", "workspace", "folder node", "folder_note", "obsidian", "page tree", "pkwf"):
        assert forbidden not in serialized


def test_offline_validator_accepts_document_node_manifest_relation() -> None:
    validator = _load_validator_module()
    contracts_root = Path(__file__).resolve().parents[2] / "contracts"
    payload = validator_json(contracts_root / "fixtures" / "valid" / "artifact-bundle.gongwen.json")
    payload["artifacts"].append(
        {
            "artifact_id": "resource.document-node-manifest",
            "kind": "resource",
            "locator": "documents/docwen-node.json",
            "logical_path": "notice/docwen-node.json",
            "suggested_name": "docwen-node.json",
            "media_type": "application/vnd.docwen.document-node+json",
            "size_bytes": 3,
            "sha256": "4444444444444444444444444444444444444444444444444444444444444444",
        }
    )
    payload["relations"].append(
        {
            "type": "resource_of",
            "source_artifact_id": "resource.document-node-manifest",
            "target_artifact_id": "document.main",
            "role": "manifest",
            "ordinal": 0,
        }
    )

    validator.validate_bundle(payload)


def test_physical_page_capability_declares_closed_relation_payloads() -> None:
    contracts_root = Path(__file__).resolve().parents[2] / "contracts"
    payload = validator_json(contracts_root / "fixtures" / "valid" / "machine.capability-list.response.json")
    capability = payload["result"]["capabilities"][0]

    assert capability["capability_id"] == "convert.pdf.to_markdown"
    assert capability["output_shape"]["relation_payloads"] == ["page_fragment", "page_resource"]
    assert capability["options_schema"]["required"] == []
    assert capability["options_schema"]["additionalProperties"] is False
    assert capability["options_schema"]["properties"]["recognize_text"] == {
        "type": "boolean",
        "default": False,
    }
    assert capability["options_schema"]["properties"]["preserve_resources"] == {
        "type": "boolean",
        "default": True,
    }
    assert "to_md_enable_ocr" not in capability["options_schema"]["properties"]
    assert "to_md_keep_images" not in capability["options_schema"]["properties"]
    assert [item["code"] for item in capability["limitations"]] == [
        "physical_page_ocr.best_effort",
        "physical_page_ocr.consumer_owned_import",
        "runtime_route_limitation",
    ]
    serialized = json.dumps(capability, ensure_ascii=False).casefold()
    for forbidden in ("page_nodes", "pkwf", "wenleaf"):
        assert forbidden not in serialized


def test_physical_page_lifecycle_trace_uses_only_v4_public_fidelity_options() -> None:
    contracts_root = Path(__file__).resolve().parents[2] / "contracts"
    trace = validator_json(contracts_root / "fixtures" / "valid" / "machine.task-lifecycle.trace.json")
    plan_request = next(message for message in trace if message.get("method") == "task/plan")
    plan_response = next(message for message in trace if message.get("id") == 3 and "result" in message)

    assert plan_request["params"]["options"]["recognize_text"] is True
    assert plan_request["params"]["options"]["preserve_resources"] is False
    assert plan_response["result"]["effective_options"]["recognize_text"] is True
    assert plan_response["result"]["effective_options"]["preserve_resources"] is False
    serialized = json.dumps([plan_request, plan_response], ensure_ascii=False)
    assert "to_md_enable_ocr" not in serialized
    assert "to_md_keep_images" not in serialized


def test_machine_schema_rejects_unknown_relation_payload_name() -> None:
    validator = _load_validator_module()
    contracts_root = Path(__file__).resolve().parents[2] / "contracts"
    payload = validator_json(contracts_root / "fixtures" / "valid" / "machine.capability-list.response.json")
    payload["result"]["capabilities"][0]["output_shape"]["relation_payloads"].append("node_tree")

    manifest = validator_json(contracts_root / "conformance-manifest.json")
    validators, _ = validator._build_validators(contracts_root, manifest)
    with pytest.raises(ValidationError):
        validators["machine"].validate(payload)


def test_offline_validator_rejects_page_fragments_owned_by_nonentry_document() -> None:
    validator = _load_validator_module()
    contracts_root = Path(__file__).resolve().parents[2] / "contracts"
    payload = validator_json(contracts_root / "fixtures" / "valid" / "artifact-bundle.ocr.json")
    primary_id = payload["entries"][0]["artifact_id"]
    secondary_id = "document.secondary"
    payload["artifacts"].append(
        {
            "artifact_id": secondary_id,
            "kind": "document",
            "locator": "secondary.md",
            "suggested_name": "secondary.md",
            "media_type": "text/markdown",
            "size_bytes": 0,
            "sha256": "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
        }
    )
    payload["relations"].append(
        {
            "type": "attachment_of",
            "source_artifact_id": secondary_id,
            "target_artifact_id": primary_id,
            "role": "attachment",
            "ordinal": 0,
        }
    )
    for relation in payload["relations"]:
        if relation["type"] == "fragment_of" and relation["role"] == "ocr_page":
            relation["target_artifact_id"] = secondary_id

    with pytest.raises(validator.SemanticContractError) as rejected:
        validator.validate_bundle(payload)

    assert rejected.value.code == "unexpected_page_semantics"


def test_offline_validator_rejects_page_resource_owned_by_nonentry_document() -> None:
    validator = _load_validator_module()
    contracts_root = Path(__file__).resolve().parents[2] / "contracts"
    payload = validator_json(contracts_root / "fixtures" / "valid" / "artifact-bundle.ocr.json")
    primary_id = payload["entries"][0]["artifact_id"]
    secondary_id = "document.secondary"
    payload["artifacts"].append(
        {
            "artifact_id": secondary_id,
            "kind": "document",
            "locator": "secondary.md",
            "suggested_name": "secondary.md",
            "media_type": "text/markdown",
            "size_bytes": 0,
            "sha256": "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
        }
    )
    payload["relations"].append(
        {
            "type": "attachment_of",
            "source_artifact_id": secondary_id,
            "target_artifact_id": primary_id,
            "role": "attachment",
            "ordinal": 0,
        }
    )
    unresolved = next(
        relation
        for relation in payload["relations"]
        if relation["type"] == "resource_of" and "page_resource" not in relation
    )
    unresolved["target_artifact_id"] = secondary_id
    unresolved["page_resource"] = {"source_page": 1}

    with pytest.raises(validator.SemanticContractError) as rejected:
        validator.validate_bundle(payload)

    assert rejected.value.code == "resource_page_mismatch"


def test_typed_input_trace_rejects_in_stable_first_error_order() -> None:
    validator = _load_validator_module()
    contracts_root = Path(__file__).resolve().parents[2] / "contracts"

    expected = [
        ("machine.duplicate-input-logical-path.trace.json", "duplicate_input_logical_path"),
        ("machine.undeclared-input-role.trace.json", "undeclared_input_role"),
        ("machine.input-slot-kind-mismatch.trace.json", "input_slot_kind_mismatch"),
        ("machine.input-slot-media-type-mismatch.trace.json", "input_slot_media_type_mismatch"),
        ("machine.input-slot-cardinality-mismatch.trace.json", "input_slot_cardinality_mismatch"),
    ]
    for filename, code in expected:
        trace = validator_json(contracts_root / "fixtures" / "invalid" / filename)
        with pytest.raises(validator.SemanticContractError) as rejected:
            validator.validate_trace(trace, requires_terminal=False)
        assert rejected.value.code == code


def test_failed_trace_rejects_artifact_bound_diagnostic_without_bundle() -> None:
    validator = _load_validator_module()
    contracts_root = Path(__file__).resolve().parents[2] / "contracts"
    trace = validator_json(contracts_root / "fixtures" / "valid" / "machine.task-failed.trace.json")
    terminal = next(message for message in trace if message.get("method") == "task/failed")
    terminal["params"]["diagnostics"] = [
        {
            "severity": "error",
            "code": "recognition_failed",
            "message": "page failed",
            "artifact_id": "artifact.unavailable",
        }
    ]

    with pytest.raises(validator.SemanticContractError) as rejected:
        validator.validate_trace(trace, requires_terminal=True)

    assert rejected.value.code == "dangling_diagnostic_artifact"


def validator_json(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))
