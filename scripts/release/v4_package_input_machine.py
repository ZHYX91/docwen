from __future__ import annotations

import base64
import binascii
import hashlib
import json
import re
from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any, cast

from scripts.release import v4_evidence_contract as evidence_contract
from scripts.release import v4_evidence_io as evidence_io

HARNESS_ID = "docwen-v4-exact-two-numbering"
HARNESS_VERSION = 1
TRANSCRIPT_SCHEMA = "docwen.v4_exact_two_stdio_transcript.v1"
EXACT_CAPABILITY_ID = "convert.markdown.to_docx"
NEUTRAL_MEDIA_TYPE = "application/vnd.docwen.resolved-document+json"
PLAN_MEDIA_TYPE = "application/vnd.docwen.numbering-export-plan+json"
DOCX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

_HEX64 = re.compile(r"^[0-9a-f]{64}$")


class V4MachineProofError(RuntimeError):
    """A Machine exact-two request or sealed transcript is not proved."""


def _json_bytes(value: object) -> bytes:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"), sort_keys=True).encode("utf-8")


def _sha256(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def _object(value: object, label: str) -> dict[str, Any]:
    if not isinstance(value, dict) or not all(isinstance(key, str) for key in value):
        raise V4MachineProofError(label)
    return cast(dict[str, Any], value)


def _array(value: object, label: str) -> list[object]:
    if not isinstance(value, list):
        raise V4MachineProofError(label)
    return cast(list[object], value)


def _reject_duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise V4MachineProofError(f"machine_transcript_duplicate_json_key:{key}")
        result[key] = value
    return result


def _json_from_bytes(raw: bytes, *, label: str) -> dict[str, Any]:
    try:
        value = json.loads(
            raw.decode("utf-8", errors="strict"),
            object_pairs_hook=_reject_duplicate_keys,
        )
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise V4MachineProofError(f"{label}_invalid_json") from exc
    return _object(value, f"{label}_not_object")


def _stable_payload(path: Path, *, label: str) -> bytes:
    try:
        _, payload, _, _ = evidence_io._read_stable(path, label=label, collect=True)
    except evidence_contract.V4EvidenceContractError as exc:
        raise V4MachineProofError(str(exc)) from exc
    assert payload is not None
    return payload


def exact_two_inputs(neutral: Path, plan: Path) -> list[dict[str, object]]:
    neutral_bytes = _stable_payload(neutral, label="neutral_document_input")
    plan_bytes = _stable_payload(plan, label="numbering_export_plan_input")
    return [
        {
            "input_id": "input.neutral-document",
            "kind": "document",
            "role": "neutral_document",
            "logical_path": "inputs/document.resolved.json",
            "locator": {"kind": "local_path", "path": str(neutral.resolve(strict=True))},
            "media_type": NEUTRAL_MEDIA_TYPE,
            "size_bytes": len(neutral_bytes),
            "sha256": _sha256(neutral_bytes),
        },
        {
            "input_id": "input.numbering-export-plan",
            "kind": "resource",
            "role": "numbering_export_plan",
            "logical_path": "inputs/numbering-export-plan.json",
            "locator": {"kind": "local_path", "path": str(plan.resolve(strict=True))},
            "media_type": PLAN_MEDIA_TYPE,
            "size_bytes": len(plan_bytes),
            "sha256": _sha256(plan_bytes),
        },
    ]


def validate_exact_two_request(inputs: object, options: object) -> None:
    if not isinstance(inputs, list) or len(inputs) != 2 or options != {}:
        raise V4MachineProofError("exact_two_input_cardinality_or_options_invalid")
    expected = (
        (
            "input.neutral-document",
            "neutral_document",
            "document",
            "inputs/document.resolved.json",
            NEUTRAL_MEDIA_TYPE,
        ),
        (
            "input.numbering-export-plan",
            "numbering_export_plan",
            "resource",
            "inputs/numbering-export-plan.json",
            PLAN_MEDIA_TYPE,
        ),
    )
    required_keys = {
        "input_id",
        "kind",
        "role",
        "logical_path",
        "locator",
        "media_type",
        "size_bytes",
        "sha256",
    }
    for raw, (input_id, role, kind, logical_path, media_type) in zip(inputs, expected, strict=True):
        if (
            not isinstance(raw, dict)
            or set(raw) != required_keys
            or raw.get("input_id") != input_id
            or raw.get("role") != role
            or raw.get("kind") != kind
            or raw.get("logical_path") != logical_path
        ):
            raise V4MachineProofError("exact_two_input_role_or_kind_invalid")
        if raw.get("media_type") != media_type:
            raise V4MachineProofError("exact_two_input_media_type_invalid")
        locator = raw.get("locator")
        if (
            not isinstance(locator, dict)
            or set(locator) != {"kind", "path"}
            or locator.get("kind") != "local_path"
            or not isinstance(locator.get("path"), str)
            or not isinstance(raw.get("size_bytes"), int)
            or isinstance(raw.get("size_bytes"), bool)
            or cast(int, raw["size_bytes"]) < 1
            or not isinstance(raw.get("sha256"), str)
            or _HEX64.fullmatch(cast(str, raw["sha256"])) is None
        ):
            raise V4MachineProofError("exact_two_input_byte_identity_invalid")
        payload = _stable_payload(Path(cast(str, locator["path"])), label=role)
        if len(payload) != raw["size_bytes"] or _sha256(payload) != raw["sha256"]:
            raise V4MachineProofError("exact_two_input_pointer_identity_mismatch")
    serialized = json.dumps({"inputs": inputs, "options": options}, ensure_ascii=False, sort_keys=True).casefold()
    if '"role": "source"' in serialized or '"role": "bibliography"' in serialized:
        raise V4MachineProofError("legacy_source_bibliography_input_rejected")


def _normalize_message(value: object) -> object:
    if isinstance(value, list):
        return [_normalize_message(item) for item in value]
    if not isinstance(value, dict):
        return value
    normalized: dict[str, object] = {}
    local_path = value.get("kind") == "local_path" and "path" in value
    for key, item in value.items():
        normalized[key] = "<local-path>" if local_path and key == "path" else _normalize_message(item)
    return normalized


def _frame(value: object) -> bytes:
    body = json.dumps(value, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    return f"Content-Length: {len(body)}\r\n\r\n".encode("ascii") + body


def transcript_event(
    *,
    direction: str,
    operation: str,
    case_id: str | None,
    raw_frame: bytes,
    message: Mapping[str, object],
) -> dict[str, object]:
    if direction not in {"request", "response"} or not operation:
        raise V4MachineProofError("machine_transcript_event_identity_invalid")
    normalized = _normalize_message(dict(message))
    if _normalize_message(_decode_frame(raw_frame)) != normalized:
        raise V4MachineProofError("machine_transcript_wire_message_mismatch")
    sealed_frame = raw_frame if direction == "response" else _frame(normalized)
    return {
        "direction": direction,
        "operation": operation,
        "caseId": case_id,
        "frameBytes": len(sealed_frame),
        "frameSha256": _sha256(sealed_frame),
        "rawFrameBase64": base64.b64encode(sealed_frame).decode("ascii"),
        "message": normalized,
    }


def transcript_request_digest(events: Sequence[Mapping[str, object]]) -> str:
    frames: list[bytes] = []
    for item in events:
        if item.get("direction") != "request":
            continue
        encoded = item.get("rawFrameBase64")
        if not isinstance(encoded, str):
            raise V4MachineProofError("machine_transcript_request_frame_missing")
        try:
            frames.append(base64.b64decode(encoded, validate=True))
        except (ValueError, binascii.Error) as exc:
            raise V4MachineProofError("machine_transcript_request_base64_invalid") from exc
    return _sha256(b"".join(frames))


def build_session_transcript(
    *,
    harness: Any,
    events: Sequence[Mapping[str, object]],
    request_digest: str,
    stdout: bytes,
    stderr: bytes,
    exit_code: int,
) -> bytes:
    event_values = [dict(item) for item in events]
    if request_digest != transcript_request_digest(event_values):
        raise V4MachineProofError("machine_transcript_request_digest_mismatch")
    payload = {
        "schema": TRANSCRIPT_SCHEMA,
        "harness": {
            "id": HARNESS_ID,
            "version": HARNESS_VERSION,
            "manifest": harness.manifest_identity,
            "executedCaseIds": [case.case_id for case in harness.cases],
        },
        "process": {"argv": ["DocWenCLI.exe", "serve", "--stdio"], "sessionCount": 1, "exitCode": exit_code},
        "streams": {
            "requestSha256": request_digest,
            "requestSetSha256": _sha256(
                _json_bytes([item["frameSha256"] for item in event_values if item.get("direction") == "request"])
            ),
            "stdoutSha256": _sha256(stdout),
            "stdoutBytes": len(stdout),
            "stderrSha256": _sha256(stderr),
            "stderrBytes": len(stderr),
        },
        "events": [{"ordinal": index, **item} for index, item in enumerate(event_values)],
    }
    raw = _json_bytes(payload)
    validate_session_transcript(raw, harness=harness)
    return raw


def _decode_frame(raw: bytes) -> dict[str, Any]:
    header, separator, body = raw.partition(b"\r\n\r\n")
    match = re.fullmatch(rb"Content-Length: ([1-9][0-9]*)", header)
    if not separator or match is None or len(body) != int(match.group(1)):
        raise V4MachineProofError("machine_transcript_frame_invalid")
    return _json_from_bytes(body, label="machine_transcript_frame")


def _forward_inputs(event: Mapping[str, object], case: Any) -> None:
    message = _object(event.get("message"), "machine_transcript_request_invalid")
    params = _object(message.get("params"), "machine_transcript_params_invalid")
    if (
        message.get("jsonrpc") != "2.0"
        or message.get("method") != "task/plan"
        or params.get("capability_id") != EXACT_CAPABILITY_ID
        or params.get("options") != {}
    ):
        raise V4MachineProofError(f"machine_transcript_forward_request_invalid:{case.case_id}")
    inputs = params.get("inputs")
    if not isinstance(inputs, list) or len(inputs) != 2:
        raise V4MachineProofError(f"machine_transcript_forward_cardinality_invalid:{case.case_id}")
    expected = (
        (
            "input.neutral-document",
            "document",
            "neutral_document",
            "inputs/document.resolved.json",
            NEUTRAL_MEDIA_TYPE,
            case.neutral_identity,
        ),
        (
            "input.numbering-export-plan",
            "resource",
            "numbering_export_plan",
            "inputs/numbering-export-plan.json",
            PLAN_MEDIA_TYPE,
            case.plan_identity,
        ),
    )
    for raw, (input_id, kind, role, logical_path, media_type, identity) in zip(inputs, expected, strict=True):
        item = _object(raw, "machine_transcript_forward_input_invalid")
        if set(item) != {
            "input_id",
            "kind",
            "role",
            "logical_path",
            "locator",
            "media_type",
            "size_bytes",
            "sha256",
        }:
            raise V4MachineProofError(f"machine_transcript_forward_input_not_closed:{case.case_id}")
        if (
            item.get("input_id") != input_id
            or item.get("kind") != kind
            or item.get("role") != role
            or item.get("logical_path") != logical_path
            or item.get("media_type") != media_type
            or item.get("size_bytes") != identity["bytes"]
            or item.get("sha256") != identity["sha256"]
            or item.get("locator") != {"kind": "local_path", "path": "<local-path>"}
        ):
            raise V4MachineProofError(f"machine_transcript_forward_input_mismatch:{case.case_id}:{role}")
    if params.get("output") != {
        "staging_root": {"kind": "local_path", "path": "<local-path>"},
        "staging_policy": "require_empty",
    }:
        raise V4MachineProofError(f"machine_transcript_forward_output_invalid:{case.case_id}")


def _rpc_result(
    request_events: Mapping[str, Mapping[str, object]],
    response_events: Mapping[str, Sequence[Mapping[str, object]]],
    operation: str,
) -> tuple[dict[str, Any], dict[str, Any]]:
    request = _object(request_events[operation].get("message"), "machine_transcript_rpc_request_invalid")
    request_id = request.get("id")
    if request.get("jsonrpc") != "2.0" or not isinstance(request_id, int) or isinstance(request_id, bool):
        raise V4MachineProofError(f"machine_transcript_rpc_request_invalid:{operation}")
    events = response_events.get(operation, ())
    matches = [
        _object(item.get("message"), "machine_transcript_rpc_response_invalid")
        for item in events
        if cast(Mapping[str, object], item.get("message", {})).get("id") == request_id
    ]
    if len(matches) != 1 or matches[0].get("jsonrpc") != "2.0" or not isinstance(matches[0].get("result"), dict):
        raise V4MachineProofError(f"machine_transcript_rpc_response_invalid:{operation}")
    if len(events) != 1:
        raise V4MachineProofError(f"machine_transcript_rpc_response_cardinality_invalid:{operation}")
    return request, cast(dict[str, Any], matches[0]["result"])


def _prove_task_exchange(
    request_events: Mapping[str, Mapping[str, object]],
    response_events: Mapping[str, Sequence[Mapping[str, object]]],
    *,
    prefix: str,
) -> None:
    kind, separator, case_id = prefix.partition(":")
    suffix = f":{case_id}" if separator else ""
    plan_operation = f"{kind}-plan{suffix}"
    execute_operation = f"{kind}-execute{suffix}"
    terminal_operation = f"{kind}-terminal{suffix}"
    plan_request, plan_result = _rpc_result(request_events, response_events, plan_operation)
    execute_request, execute_result = _rpc_result(request_events, response_events, execute_operation)
    plan_id = plan_result.get("plan_id")
    task_id = execute_result.get("task_id")
    if (
        plan_request.get("method") != "task/plan"
        or execute_request.get("method") != "task/execute"
        or not isinstance(plan_id, str)
        or not plan_id
        or execute_request.get("params") != {"plan_id": plan_id}
        or not isinstance(task_id, str)
        or not task_id
        or execute_result.get("state") != "accepted"
    ):
        raise V4MachineProofError(f"machine_transcript_task_exchange_invalid:{prefix}")
    terminal_events = response_events.get(terminal_operation, ())
    messages = [_object(item.get("message"), "machine_transcript_terminal_invalid") for item in terminal_events]
    terminal_methods = {"task/completed", "task/failed", "task/cancelled"}
    if any(item.get("method") not in {*terminal_methods, "task/progress"} for item in messages):
        raise V4MachineProofError(f"machine_transcript_terminal_event_invalid:{prefix}")
    for item in messages:
        if (
            item.get("method") == "task/progress"
            and _object(item.get("params"), "machine_transcript_progress_params_invalid").get("task_id") != task_id
        ):
            raise V4MachineProofError(f"machine_transcript_progress_task_invalid:{prefix}")
    terminals = [item for item in messages if item.get("method") in terminal_methods]
    if (
        len(terminals) != 1
        or terminals[0].get("method") != "task/completed"
        or _object(terminals[0].get("params"), "machine_transcript_terminal_params_invalid").get("task_id") != task_id
    ):
        raise V4MachineProofError(f"machine_transcript_terminal_invalid:{prefix}")


def validate_session_transcript(
    raw: bytes,
    *,
    harness: Any,
    outputs: Sequence[Any] | None = None,
) -> dict[str, Any]:
    value = _json_from_bytes(raw, label="machine_transcript")
    if set(value) != {"schema", "harness", "process", "streams", "events"} or value.get("schema") != TRANSCRIPT_SCHEMA:
        raise V4MachineProofError("machine_transcript_not_closed")
    if _object(value.get("harness"), "machine_transcript_harness_invalid") != {
        "id": HARNESS_ID,
        "version": HARNESS_VERSION,
        "manifest": harness.manifest_identity,
        "executedCaseIds": [case.case_id for case in harness.cases],
    }:
        raise V4MachineProofError("machine_transcript_harness_mismatch")
    if value.get("process") != {"argv": ["DocWenCLI.exe", "serve", "--stdio"], "sessionCount": 1, "exitCode": 0}:
        raise V4MachineProofError("machine_transcript_process_invalid")
    streams = _object(value.get("streams"), "machine_transcript_streams_invalid")
    if set(streams) != {
        "requestSha256",
        "requestSetSha256",
        "stdoutSha256",
        "stdoutBytes",
        "stderrSha256",
        "stderrBytes",
    }:
        raise V4MachineProofError("machine_transcript_streams_not_closed")
    if (
        any(
            _HEX64.fullmatch(str(streams.get(key, ""))) is None
            for key in ("requestSha256", "requestSetSha256", "stdoutSha256", "stderrSha256")
        )
        or streams.get("stderrBytes") != 0
        or streams.get("stderrSha256") != _sha256(b"")
    ):
        raise V4MachineProofError("machine_transcript_stream_identity_invalid")
    events = _array(value.get("events"), "machine_transcript_events_invalid")
    request_events: list[dict[str, Any]] = []
    request_frames: list[bytes] = []
    response_frames: list[bytes] = []
    response_events: dict[str, list[dict[str, Any]]] = {}
    event_keys = {
        "ordinal",
        "direction",
        "operation",
        "caseId",
        "frameBytes",
        "frameSha256",
        "rawFrameBase64",
        "message",
    }
    for ordinal, raw_event in enumerate(events):
        event = _object(raw_event, "machine_transcript_event_invalid")
        if (
            set(event) != event_keys
            or event.get("ordinal") != ordinal
            or event.get("direction") not in {"request", "response"}
            or not isinstance(event.get("operation"), str)
            or not isinstance(event.get("frameBytes"), int)
            or isinstance(event.get("frameBytes"), bool)
            or cast(int, event["frameBytes"]) < 1
            or _HEX64.fullmatch(str(event.get("frameSha256", ""))) is None
            or not isinstance(event.get("message"), dict)
        ):
            raise V4MachineProofError("machine_transcript_event_not_closed")
        encoded = event.get("rawFrameBase64")
        if not isinstance(encoded, str):
            raise V4MachineProofError("machine_transcript_frame_missing")
        try:
            frame = base64.b64decode(encoded, validate=True)
        except (ValueError, binascii.Error) as exc:
            raise V4MachineProofError("machine_transcript_base64_invalid") from exc
        if len(frame) != event["frameBytes"] or _sha256(frame) != event["frameSha256"]:
            raise V4MachineProofError("machine_transcript_frame_identity_invalid")
        if _normalize_message(_decode_frame(frame)) != event["message"]:
            raise V4MachineProofError("machine_transcript_message_mismatch")
        if event["direction"] == "request":
            if b"<local-path>" not in frame and b'"kind":"local_path"' in frame:
                raise V4MachineProofError("machine_transcript_request_leaks_raw_path")
            request_events.append(event)
            request_frames.append(frame)
        else:
            response_frames.append(frame)
            operation = str(event["operation"])
            response_events.setdefault(operation, []).append(event)
    expected_operations = ["initialize", "discovery", "validation-plan", "validation-execute"]
    for case in harness.cases:
        expected_operations.extend((f"forward-plan:{case.case_id}", f"forward-execute:{case.case_id}"))
        expected_operations.extend((f"reverse-plan:{case.case_id}", f"reverse-execute:{case.case_id}"))
    if [str(item["operation"]) for item in request_events] != expected_operations:
        raise V4MachineProofError("machine_transcript_request_order_invalid")
    terminal_operations = {"validation-terminal"}
    for case in harness.cases:
        terminal_operations.update((f"forward-terminal:{case.case_id}", f"reverse-terminal:{case.case_id}"))
    if not set(response_events) <= set(expected_operations) | terminal_operations:
        raise V4MachineProofError("machine_transcript_response_operation_invalid")
    request_hashes = [item["frameSha256"] for item in request_events]
    if streams["requestSetSha256"] != _sha256(_json_bytes(request_hashes)):
        raise V4MachineProofError("machine_transcript_request_set_hash_invalid")
    if streams["requestSha256"] != _sha256(b"".join(request_frames)):
        raise V4MachineProofError("machine_transcript_request_digest_invalid")
    stdout = b"".join(response_frames)
    if streams["stdoutBytes"] != len(stdout) or streams["stdoutSha256"] != _sha256(stdout):
        raise V4MachineProofError("machine_transcript_stdout_identity_invalid")
    output_by_case = {item.case_id: item for item in outputs or ()}
    if outputs is not None and (
        list(output_by_case) != [case.case_id for case in harness.cases] or len(output_by_case) != len(outputs)
    ):
        raise V4MachineProofError("machine_transcript_output_case_set_invalid")
    by_operation = {str(item["operation"]): item for item in request_events}
    request_ids = [
        _object(item.get("message"), "machine_transcript_request_invalid").get("id") for item in request_events
    ]
    if any(not isinstance(item, int) or isinstance(item, bool) or item < 1 for item in request_ids) or len(
        set(request_ids)
    ) != len(request_ids):
        raise V4MachineProofError("machine_transcript_request_id_invalid")
    initialize, _initialize_result = _rpc_result(by_operation, response_events, "initialize")
    discovery, _discovery_result = _rpc_result(by_operation, response_events, "discovery")
    if initialize.get("method") != "initialize" or discovery.get("method") != "capability/list":
        raise V4MachineProofError("machine_transcript_session_bootstrap_invalid")
    _prove_task_exchange(by_operation, response_events, prefix="validation")
    validation = _object(
        _object(by_operation["validation-plan"].get("message"), "machine_transcript_validation_invalid").get("params"),
        "machine_transcript_validation_params_invalid",
    )
    validation_inputs = validation.get("inputs")
    if (
        validation.get("capability_id") != "validate.markdown"
        or validation.get("options") != {}
        or not isinstance(validation_inputs, list)
        or len(validation_inputs) != 1
        or _object(validation_inputs[0], "machine_transcript_validation_input_invalid").get("role") != "source"
    ):
        raise V4MachineProofError("machine_transcript_validation_request_invalid")
    for case in harness.cases:
        _forward_inputs(by_operation[f"forward-plan:{case.case_id}"], case)
        _prove_task_exchange(by_operation, response_events, prefix=f"forward:{case.case_id}")
        _prove_task_exchange(by_operation, response_events, prefix=f"reverse:{case.case_id}")
        reverse = _object(
            by_operation[f"reverse-plan:{case.case_id}"].get("message"),
            "machine_transcript_reverse_invalid",
        )
        params = _object(reverse.get("params"), "machine_transcript_reverse_params_invalid")
        reverse_inputs = params.get("inputs")
        if (
            reverse.get("method") != "task/plan"
            or params.get("capability_id") != "convert.docx.to_markdown"
            or params.get("options") != {}
            or not isinstance(reverse_inputs, list)
            or len(reverse_inputs) != 1
        ):
            raise V4MachineProofError(f"machine_transcript_reverse_request_invalid:{case.case_id}")
        reverse_input = _object(reverse_inputs[0], "machine_transcript_reverse_input_invalid")
        if (
            reverse_input.get("role") != "source"
            or reverse_input.get("media_type") != DOCX_MEDIA_TYPE
            or reverse_input.get("locator") != {"kind": "local_path", "path": "<local-path>"}
        ):
            raise V4MachineProofError(f"machine_transcript_reverse_input_invalid:{case.case_id}")
        if outputs is not None:
            docx = output_by_case[case.case_id].docx
            if reverse_input.get("size_bytes") != len(docx) or reverse_input.get("sha256") != _sha256(docx):
                raise V4MachineProofError(f"machine_transcript_reverse_docx_mismatch:{case.case_id}")
    return value
