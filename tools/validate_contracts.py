"""Validate DocWen-owned wire contracts and their conformance fixtures."""

from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, cast

from jsonschema import Draft202012Validator
from jsonschema.exceptions import ValidationError
from referencing import Registry, Resource

REPO_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_CONTRACTS_ROOT = REPO_ROOT / "contracts"
TERMINAL_METHODS = frozenset({"task/completed", "task/failed", "task/cancelled"})
TASK_NOTIFICATION_METHODS = frozenset({"task/progress", *TERMINAL_METHODS})
STRUCTURAL_RELATIONS = frozenset({"attachment_of", "fragment_of", "resource_of"})
PAGE_OCR_STATUSES = frozenset(
    {
        "success",
        "no_text",
        "input_missing",
        "unavailable",
        "model_missing",
        "initialization_failed",
        "recognition_failed",
    }
)
MAX_MESSAGE_BYTES = 16 * 1024 * 1024
FRAME_HEADER = re.compile(rb"Content-Length: ([1-9][0-9]*)")
LOGICAL_INPUT_URI = re.compile(r"^[A-Za-z][A-Za-z0-9+.-]*:")


class SemanticContractError(ValueError):
    """A cross-field or graph invariant not expressible in portable JSON Schema."""

    def __init__(self, code: str, message: str) -> None:
        super().__init__(f"{code}: {message}")
        self.code = code


@dataclass(frozen=True)
class ValidationSummary:
    schemas: int
    valid_fixtures: int
    invalid_fixtures: int


def _load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def _fail(code: str, message: str) -> None:
    raise SemanticContractError(code, message)


def _validate_relative_locator(locator: str) -> None:
    segments = locator.split("/")
    if (
        "\\" in locator
        or locator.startswith("/")
        or (len(locator) >= 2 and locator[0].isalpha() and locator[1] == ":")
        or any(segment in {"", ".", ".."} for segment in segments)
    ):
        _fail("invalid_artifact_locator", f"artifact locator must be a normalized relative POSIX path: {locator!r}")


def _is_positive_int(value: object) -> bool:
    return isinstance(value, int) and not isinstance(value, bool) and value > 0


def _validate_input_logical_path(logical_path: str) -> None:
    """Reject non-key input paths before any capability slot comparison."""
    if (
        not logical_path
        or logical_path.startswith("/")
        or "\\" in logical_path
        or "\x00" in logical_path
        or LOGICAL_INPUT_URI.match(logical_path)
    ):
        _fail("invalid_input_logical_path", f"input logical_path is not a relative POSIX key: {logical_path!r}")
    if any(segment in {"", ".", ".."} for segment in logical_path.split("/")):
        _fail("invalid_input_logical_path", f"input logical_path has an invalid segment: {logical_path!r}")


def _validate_task_plan_inputs(
    params: dict[str, Any],
    capabilities: dict[str, dict[str, Any]],
) -> None:
    """Validate typed plan inputs in stable, caller-actionable error order."""
    inputs = params["inputs"]
    input_ids = [input_handle["input_id"] for input_handle in inputs]
    if len(input_ids) != len(set(input_ids)):
        _fail("duplicate_input_id", "task/plan input_id values must be unique")

    for input_handle in inputs:
        _validate_input_logical_path(input_handle["logical_path"])

    logical_paths = [input_handle["logical_path"] for input_handle in inputs]
    if len(logical_paths) != len(set(logical_paths)):
        _fail("duplicate_input_logical_path", "task/plan logical_path values must be globally unique")

    capability = capabilities.get(params["capability_id"])
    if capability is None:
        return
    slots = {slot["role"]: slot for slot in capability["input_shape"]["slots"]}

    for input_handle in inputs:
        if input_handle["role"] not in slots:
            _fail("undeclared_input_role", f"input role {input_handle['role']!r} is not declared by capability")

    for input_handle in inputs:
        slot = slots[input_handle["role"]]
        if input_handle["kind"] != slot["kind"]:
            _fail("input_slot_kind_mismatch", f"input role {input_handle['role']!r} has incompatible kind")

    for input_handle in inputs:
        slot = slots[input_handle["role"]]
        if input_handle["media_type"] not in slot["media_types"]:
            _fail("input_slot_media_type_mismatch", f"input role {input_handle['role']!r} has incompatible media type")

    for role, slot in slots.items():
        count = sum(input_handle["role"] == role for input_handle in inputs)
        maximum = slot.get("max_items")
        if count < slot["min_items"] or (maximum is not None and count > maximum):
            _fail("input_slot_cardinality_mismatch", f"input role {role!r} violates slot cardinality")


def _validate_capability_list(capabilities: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    """Enforce D2 slot invariants that portable JSON Schema cannot express."""
    result: dict[str, dict[str, Any]] = {}
    for capability in capabilities:
        slots = capability["input_shape"]["slots"]
        roles = [slot["role"] for slot in slots]
        if len(roles) != len(set(roles)):
            _fail("duplicate_input_slot_role", "capability input_shape slot roles must be unique")
        is_resolved_markdown = capability["capability_id"] == "convert.markdown.to_docx"
        if is_resolved_markdown:
            expected = {
                "neutral_document": (
                    "document",
                    "application/vnd.docwen.resolved-document+json",
                ),
                "numbering_export_plan": (
                    "resource",
                    "application/vnd.docwen.numbering-export-plan+json",
                ),
            }
            if set(roles) != set(expected):
                _fail(
                    "invalid_resolved_numbering_slots",
                    "convert.markdown.to_docx requires the exact neutral-document/plan pair",
                )
            for slot in slots:
                expected_kind, expected_media = expected[slot["role"]]
                if (
                    slot["kind"] != expected_kind
                    or slot["media_types"] != [expected_media]
                    or slot["min_items"] != 1
                    or slot.get("max_items") != 1
                ):
                    _fail(
                        "invalid_resolved_numbering_slots",
                        "resolved-numbering slots have non-canonical kind/media/cardinality",
                    )
            properties = capability["options_schema"].get("properties", {})
            legacy_numbering = {
                "remove_numbering",
                "add_numbering",
                "numbering_scheme",
                "heading_numbering_render_mode",
            }
            if legacy_numbering.intersection(properties):
                _fail(
                    "legacy_numbering_option_exposed",
                    "resolved-numbering capability may not expose superseded numbering controls",
                )
        elif not any(slot["role"] == "source" and slot["min_items"] >= 1 for slot in slots):
            _fail("missing_required_source_slot", "capability input_shape requires source min_items >= 1")
        for slot in slots:
            if not is_resolved_markdown and slot["role"] != "source" and slot["kind"] != "resource":
                _fail("invalid_input_slot_kind", "non-source input slots must have resource kind")
            if slot.get("max_items") is not None and slot["max_items"] < slot["min_items"]:
                _fail("invalid_input_slot_cardinality", "input slot max_items must be at least min_items")
        result[capability["capability_id"]] = capability
    return result


def validate_bundle(bundle: dict[str, Any]) -> None:
    """Validate Artifact Bundle graph, ownership, ordering, and locator invariants."""

    artifacts = bundle["artifacts"]
    artifact_ids = [artifact["artifact_id"] for artifact in artifacts]
    if len(artifact_ids) != len(set(artifact_ids)):
        _fail("duplicate_artifact_id", "artifact_id values must be unique within a bundle")

    artifact_by_id = {artifact["artifact_id"]: artifact for artifact in artifacts}
    for artifact in artifacts:
        if artifact["suggested_name"] in {".", ".."}:
            _fail("invalid_suggested_name", "suggested_name must be a safe basename")

    entries = bundle["entries"]
    entry_ids = [entry["artifact_id"] for entry in entries]
    if len(entry_ids) != len(set(entry_ids)):
        _fail("duplicate_entry", "an artifact may appear in entries at most once")
    entry_ordinals = [entry["ordinal"] for entry in entries]
    if len(entry_ordinals) != len(set(entry_ordinals)):
        _fail("duplicate_entry_ordinal", "entry ordinal values must be unique")
    for entry in entries:
        artifact_id = entry["artifact_id"]
        if artifact_id not in artifact_by_id:
            _fail("dangling_entry", f"entry references missing artifact {artifact_id!r}")
        kind = artifact_by_id[artifact_id]["kind"]
        role = entry["role"]
        if role == "ocr_page" and kind != "fragment":
            _fail("incompatible_entry_role", f"entry role {role!r} requires a fragment artifact")
        if role == "section" and kind not in {"document", "fragment"}:
            _fail("incompatible_entry_role", "entry role 'section' requires a document or fragment artifact")
        if role == "image" and kind != "resource":
            _fail("incompatible_entry_role", "entry role 'image' requires a resource artifact")

    preferred_entries = [entry for entry in entries if entry["preferred"]]
    if len(preferred_entries) > 1:
        _fail("preferred_entry_count", "a bundle may have at most one preferred entry")
    primary_document_ids = {
        entry["artifact_id"]
        for entry in entries
        if entry["role"] == "primary" and artifact_by_id[entry["artifact_id"]]["kind"] == "document"
    }

    adjacency: dict[str, set[str]] = {artifact_id: set() for artifact_id in artifact_ids}
    structural_owner: dict[str, str] = {}
    ordered_relation_slots: set[tuple[str, str, int]] = set()
    relation_roles = {
        "attachment_of": {"attachment"},
        "fragment_of": {"ocr_page", "ocr_text", "section", "worksheet"},
        "resource_of": {"image", "manifest", "original", "preview", "worksheet"},
        "derived_from": {"source", "original"},
    }
    relation_kinds = {
        "attachment_of": ({"document"}, {"document"}),
        "fragment_of": ({"fragment"}, {"document"}),
        "resource_of": ({"resource"}, {"document", "fragment"}),
        "derived_from": ({"document", "fragment", "resource"}, {"document", "fragment", "resource"}),
    }

    directed_edges: dict[str, set[str]] = {artifact_id: set() for artifact_id in artifact_ids}
    page_relations: list[dict[str, Any]] = []
    for relation in bundle["relations"]:
        source = relation["source_artifact_id"]
        target = relation["target_artifact_id"]
        relation_type = relation["type"]
        if source not in artifact_by_id or target not in artifact_by_id:
            _fail("dangling_relation", f"relation {source!r} -> {target!r} references a missing artifact")
        if source == target:
            _fail("self_relation", f"artifact {source!r} cannot relate to itself")

        source_kinds, target_kinds = relation_kinds[relation_type]
        if (
            artifact_by_id[source]["kind"] not in source_kinds
            or artifact_by_id[target]["kind"] not in target_kinds
            or relation["role"] not in relation_roles[relation_type]
        ):
            _fail("incompatible_relation", f"relation {relation_type!r} has incompatible kinds or role")

        if relation_type in STRUCTURAL_RELATIONS:
            if source in structural_owner:
                _fail("multiple_structural_owners", f"artifact {source!r} has more than one structural owner")
            structural_owner[source] = target
            if source in entry_ids:
                _fail("owned_entry", f"entry artifact {source!r} cannot also have a structural owner")

        if relation_type in {"attachment_of", "fragment_of"} and "ordinal" not in relation:
            _fail("missing_relation_ordinal", f"ordered relation {relation_type!r} requires ordinal")
        is_page_fragment = relation_type == "fragment_of" and relation["role"] == "ocr_page"
        if "ordinal" in relation and not is_page_fragment:
            slot = (relation_type, target, relation["ordinal"])
            if slot in ordered_relation_slots:
                _fail("duplicate_relation_ordinal", f"duplicate ordinal for {relation_type!r} targeting {target!r}")
            ordered_relation_slots.add(slot)

        adjacency[source].add(target)
        adjacency[target].add(source)
        directed_edges[source].add(target)

    for relation in bundle["relations"]:
        is_page_fragment = relation["type"] == "fragment_of" and relation["role"] == "ocr_page"
        if is_page_fragment:
            if "page_fragment" not in relation:
                _fail("missing_page_fragment_semantics", "fragment_of/ocr_page requires page_fragment semantics")
            if relation["target_artifact_id"] not in primary_document_ids:
                _fail("unexpected_page_semantics", "physical page fragments must target a primary document entry")
            page_relations.append(relation)
        elif "page_fragment" in relation:
            _fail("unexpected_page_semantics", "page_fragment is only valid on fragment_of/ocr_page")
        resource_payload_allowed = relation["type"] == "resource_of" and relation["role"] in {
            "image",
            "original",
            "preview",
        }
        if "page_resource" in relation and not resource_payload_allowed:
            _fail("unexpected_page_semantics", "page_resource is only valid on image/original/preview resources")

    page_relation_by_artifact: dict[str, dict[str, Any]] = {}
    pages_by_owner: dict[str, list[dict[str, Any]]] = {}
    for relation in page_relations:
        page = relation["page_fragment"]
        values = (page["page_index"], page["page_count"], page["source_page"])
        if (
            page["fragment_kind"] != "page"
            or any(not _is_positive_int(value) for value in values)
            or page["page_index"] > page["page_count"]
            or page["source_page"] > page["page_count"]
            or page["ocr_status"] not in PAGE_OCR_STATUSES
        ):
            _fail("invalid_page_range", "physical page numbers and page_count must form positive in-range values")
        if relation.get("ordinal") != page["page_index"] - 1:
            _fail("page_ordinal_mismatch", "page fragment ordinal must equal page_index - 1")
        page_relation_by_artifact[relation["source_artifact_id"]] = relation
        pages_by_owner.setdefault(relation["target_artifact_id"], []).append(relation)

    for owner, owner_relations in pages_by_owner.items():
        counts = {relation["page_fragment"]["page_count"] for relation in owner_relations}
        if len(counts) != 1:
            _fail("page_count_mismatch", f"page fragments for owner {owner!r} disagree on page_count")
        page_count = next(iter(counts))
        page_indexes = [relation["page_fragment"]["page_index"] for relation in owner_relations]
        if len(page_indexes) != len(set(page_indexes)):
            _fail("duplicate_page_index", f"page fragments for owner {owner!r} repeat page_index")
        expected = set(range(1, page_count + 1))
        if len(owner_relations) != page_count or set(page_indexes) != expected:
            _fail("incomplete_page_sequence", f"page fragments for owner {owner!r} do not cover 1..page_count")
        source_pages = {relation["page_fragment"]["source_page"] for relation in owner_relations}
        if len(source_pages) != page_count or source_pages != expected:
            _fail("page_source_mismatch", f"source_page values for owner {owner!r} do not cover 1..page_count")

    for relation in bundle["relations"]:
        if relation["type"] != "resource_of" or relation["role"] not in {"image", "original", "preview"}:
            continue
        page_resource = relation.get("page_resource")
        if page_resource is not None and not _is_positive_int(page_resource["source_page"]):
            _fail("invalid_page_range", "page_resource source_page must be a positive integer")
        target_page_relation = page_relation_by_artifact.get(relation["target_artifact_id"])
        if target_page_relation is not None:
            target_page = target_page_relation["page_fragment"]
            if page_resource is None or page_resource["source_page"] != target_page["source_page"]:
                _fail("resource_page_mismatch", "page-owned resource does not match its target page fragment")
            continue
        if artifact_by_id[relation["target_artifact_id"]]["kind"] == "fragment" and page_resource is not None:
            _fail("resource_page_mismatch", "page_resource targets a fragment without physical-page semantics")
        if page_resource is not None and relation["target_artifact_id"] not in primary_document_ids:
            _fail("resource_page_mismatch", "document-owned page resource must target a primary document entry")
        if page_resource is not None and relation["target_artifact_id"] in pages_by_owner:
            _fail("resource_page_mismatch", "proven page resource must target the matching page fragment")

    visited: set[str] = set()

    def visit_for_cycle(artifact_id: str, active: set[str]) -> None:
        if artifact_id in active:
            _fail("relation_cycle", f"relation graph contains a cycle at {artifact_id!r}")
        if artifact_id in visited:
            return
        active.add(artifact_id)
        for target in directed_edges[artifact_id]:
            visit_for_cycle(target, active)
        active.remove(artifact_id)
        visited.add(artifact_id)

    for artifact_id in artifact_ids:
        visit_for_cycle(artifact_id, set())

    reachable = set(entry_ids)
    pending = list(entry_ids)
    while pending:
        artifact_id = pending.pop()
        for neighbor in adjacency[artifact_id] - reachable:
            reachable.add(neighbor)
            pending.append(neighbor)
    orphans = set(artifact_ids) - reachable
    if orphans:
        _fail("orphan_artifact", f"artifacts are not connected to any entry: {sorted(orphans)!r}")

    locators = [artifact["locator"] for artifact in artifacts]
    if len(locators) != len(set(locators)):
        _fail("duplicate_artifact_locator", "artifact locators must be unique within a bundle")
    for artifact in artifacts:
        _validate_relative_locator(artifact["locator"])


def validate_trace(messages: list[dict[str, Any]], *, requires_terminal: bool) -> None:
    """Validate request/response correlation and asynchronous task lifecycle semantics."""

    pending_requests: dict[str | int, str] = {}
    used_request_ids: set[str | int] = set()
    accepted_tasks: set[str] = set()
    terminal_tasks: set[str] = set()
    last_sequence: dict[str, int] = {}
    capabilities: dict[str, dict[str, Any]] = {}

    for message in messages:
        method = message.get("method")
        message_id = message.get("id")
        if method is not None and "id" in message:
            request_id = cast(str | int, message["id"])
            if request_id in used_request_ids:
                _fail("duplicate_request_id", f"request id {request_id!r} is reused within a trace")
            used_request_ids.add(request_id)
            pending_requests[request_id] = method
            if method == "task/plan":
                _validate_task_plan_inputs(message["params"], capabilities)
            continue

        if "result" in message or "error" in message:
            if message_id is None and "error" in message:
                continue
            response_id = cast(str | int, message_id)
            request_method = pending_requests.pop(response_id, None)
            if request_method is None:
                _fail("unmatched_response", f"response id {response_id!r} has no pending request")
            if request_method == "task/execute" and "result" in message:
                task_id = message["result"]["task_id"]
                if task_id in accepted_tasks:
                    _fail("duplicate_task_id", f"task id {task_id!r} was accepted more than once")
                accepted_tasks.add(task_id)
            if request_method == "capability/list" and "result" in message:
                capabilities = _validate_capability_list(message["result"]["capabilities"])
            if request_method == "task/cancel" and "result" in message:
                result = message["result"]
                if result["state"] != "not_found" and result["task_id"] not in accepted_tasks:
                    _fail("unknown_cancel_task", f"cancel response references unaccepted task {result['task_id']!r}")
            continue

        if method not in TASK_NOTIFICATION_METHODS:
            continue
        params = message["params"]
        task_id = params["task_id"]
        if task_id not in accepted_tasks:
            _fail("notification_before_acceptance", f"notification references unaccepted task {task_id!r}")
        if task_id in terminal_tasks:
            _fail("notification_after_terminal", f"task {task_id!r} emitted a notification after terminal state")
        sequence = params["sequence"]
        if sequence <= last_sequence.get(task_id, 0):
            _fail("nonmonotonic_sequence", f"task {task_id!r} sequence did not increase")
        last_sequence[task_id] = sequence

        if method == "task/progress" and params["completed"] > params["total"]:
            _fail("invalid_progress", f"task {task_id!r} progress exceeds total")
        if method == "task/completed":
            bundle = params["bundle"]
            if bundle["task_id"] != task_id:
                _fail("bundle_task_mismatch", f"completed bundle does not belong to task {task_id!r}")
            validate_bundle(bundle)
            artifact_ids = {artifact["artifact_id"] for artifact in bundle["artifacts"]}
            if any(
                diagnostic.get("artifact_id") not in artifact_ids
                for diagnostic in params["diagnostics"]
                if "artifact_id" in diagnostic
            ):
                _fail("dangling_diagnostic_artifact", "completed diagnostic references a missing Bundle artifact")
        elif method in {"task/failed", "task/cancelled"} and any(
            "artifact_id" in diagnostic for diagnostic in params["diagnostics"]
        ):
            _fail(
                "dangling_diagnostic_artifact",
                "failed or cancelled diagnostic cannot reference a Bundle artifact",
            )
        if method in TERMINAL_METHODS:
            terminal_tasks.add(task_id)

    if pending_requests:
        _fail("unmatched_request", f"trace ends with pending request ids: {sorted(pending_requests)!r}")
    if requires_terminal and accepted_tasks != terminal_tasks:
        missing = sorted(accepted_tasks - terminal_tasks)
        _fail("missing_terminal", f"accepted tasks have no terminal notification: {missing!r}")


def validate_framing_fixture(payload: dict[str, Any]) -> list[dict[str, Any]]:
    """Decode a chunked Content-Length stream and compare its messages with the fixture oracle."""

    stream = "".join(payload["chunks"]).encode("utf-8")
    messages: list[dict[str, Any]] = []
    cursor = 0
    while cursor < len(stream):
        header_end = stream.find(b"\r\n\r\n", cursor)
        if header_end < 0:
            _fail("invalid_frame_header", "frame header must end with CRLF CRLF")
        header = stream[cursor:header_end]
        match = FRAME_HEADER.fullmatch(header)
        if match is None:
            _fail("invalid_frame_header", "frame must contain exactly one canonical Content-Length header")
        content_length = int(match.group(1))
        if content_length > MAX_MESSAGE_BYTES:
            _fail("frame_too_large", f"frame exceeds {MAX_MESSAGE_BYTES} bytes")
        body_start = header_end + 4
        body_end = body_start + content_length
        if body_end > len(stream):
            _fail("frame_length_mismatch", "declared Content-Length exceeds available UTF-8 bytes")
        body = stream[body_start:body_end]
        try:
            message = json.loads(body.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            _fail("invalid_frame_payload", f"frame payload is not valid UTF-8 JSON: {exc}")
        if not isinstance(message, dict):
            _fail("invalid_frame_payload", "frame payload must be one JSON-RPC object")
        messages.append(message)
        cursor = body_end
    if messages != payload["expected_messages"]:
        _fail("framing_oracle_mismatch", "decoded messages do not match expected_messages")
    return messages


def _build_validators(
    contracts_root: Path,
    manifest: dict[str, Any],
) -> tuple[dict[str, Draft202012Validator], set[Path]]:
    schemas: dict[str, dict[str, Any]] = {}
    schema_paths: set[Path] = set()
    resources: list[tuple[str, Resource[Any]]] = []
    for record in manifest["schemas"]:
        path = contracts_root / record["path"]
        schema = _load_json(path)
        Draft202012Validator.check_schema(schema)
        if schema["$id"] != record["id"]:
            raise ValueError(f"schema id mismatch for {path}: {schema['$id']!r} != {record['id']!r}")
        schemas[record["name"]] = schema
        schema_paths.add(path.resolve())
        resources.append((record["id"], Resource.from_contents(schema)))
    registry = Registry().with_resources(resources)
    validators = {name: Draft202012Validator(schema, registry=registry) for name, schema in schemas.items()}
    return validators, schema_paths


def _validate_document_semantics(document_type: str, payload: Any, *, requires_terminal: bool) -> None:
    if document_type == "bundle":
        validate_bundle(payload)
    elif document_type == "trace":
        validate_trace(payload, requires_terminal=requires_terminal)


def validate_contract_set(contracts_root: Path = DEFAULT_CONTRACTS_ROOT) -> ValidationSummary:
    """Validate schemas, fixture inventory, and every expected pass/fail result."""

    manifest_path = contracts_root / "conformance-manifest.json"
    manifest = _load_json(manifest_path)
    if manifest["contract_set"] != "docwen.contracts.v1":
        raise ValueError("unexpected contract_set")
    validators, schema_paths = _build_validators(contracts_root, manifest)

    fixture_records = manifest["fixtures"]
    recorded_fixture_paths = {(contracts_root / record["path"]).resolve() for record in fixture_records}
    actual_fixture_paths = {path.resolve() for path in (contracts_root / "fixtures").rglob("*.json")}
    if recorded_fixture_paths != actual_fixture_paths:
        missing = sorted(str(path) for path in actual_fixture_paths - recorded_fixture_paths)
        stale = sorted(str(path) for path in recorded_fixture_paths - actual_fixture_paths)
        raise ValueError(f"conformance fixture inventory mismatch; missing={missing!r}; stale={stale!r}")

    actual_schema_paths = {path.resolve() for path in (contracts_root / "schemas").glob("*.json")}
    if schema_paths != actual_schema_paths:
        raise ValueError("conformance schema inventory does not match schemas directory")

    valid_count = 0
    invalid_count = 0
    for record in fixture_records:
        path = contracts_root / record["path"]
        payload = _load_json(path)
        validator = validators[record["schema"]]
        expectation = record["expect"]

        if record["document_type"] == "framing":
            try:
                documents = validate_framing_fixture(payload)
            except SemanticContractError as exc:
                if expectation != "invalid_semantic" or exc.code != record.get("error_code"):
                    raise AssertionError(
                        f"framing fixture {path} failed with unexpected semantic error {exc.code!r}"
                    ) from exc
                invalid_count += 1
                continue
            if expectation != "valid":
                raise AssertionError(f"framing fixture was expected to fail: {path}")
            for document in documents:
                validator.validate(document)
            valid_count += 1
            continue

        documents = payload if record["document_type"] == "trace" else [payload]

        if expectation == "invalid_schema":
            try:
                for document in documents:
                    validator.validate(document)
            except ValidationError:
                invalid_count += 1
                continue
            raise AssertionError(f"fixture was expected to fail schema validation: {path}")

        for document in documents:
            validator.validate(document)

        if expectation == "invalid_semantic":
            try:
                _validate_document_semantics(
                    record["document_type"],
                    payload,
                    requires_terminal=record.get("requires_terminal", False),
                )
            except SemanticContractError as exc:
                if exc.code != record["error_code"]:
                    raise AssertionError(
                        f"fixture {path} failed with {exc.code!r}; expected {record['error_code']!r}"
                    ) from exc
                invalid_count += 1
                continue
            raise AssertionError(f"fixture was expected to fail semantic validation: {path}")

        if expectation != "valid":
            raise ValueError(f"unknown fixture expectation {expectation!r} in {path}")
        _validate_document_semantics(
            record["document_type"],
            payload,
            requires_terminal=record.get("requires_terminal", False),
        )
        valid_count += 1

    return ValidationSummary(
        schemas=len(validators),
        valid_fixtures=valid_count,
        invalid_fixtures=invalid_count,
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--contracts-root",
        type=Path,
        default=DEFAULT_CONTRACTS_ROOT,
        help="Path to the DocWen contracts directory.",
    )
    args = parser.parse_args(argv)
    try:
        summary = validate_contract_set(args.contracts_root.resolve())
    except (AssertionError, OSError, ValueError) as exc:
        print(f"DocWen contract conformance failed: {exc}", file=sys.stderr)
        return 1
    print(
        "DocWen contract conformance passed: "
        f"{summary.schemas} schemas, {summary.valid_fixtures} valid fixtures, "
        f"{summary.invalid_fixtures} rejected fixtures."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
