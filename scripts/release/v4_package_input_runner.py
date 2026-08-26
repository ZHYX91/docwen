from __future__ import annotations

import hashlib
import json
import os
import subprocess
from collections.abc import Mapping
from pathlib import Path
from typing import Any

from scripts.release import v4_package_input_contract as input_contract
from scripts.release import v4_package_input_machine as machine_proof

V4PackageInputError = input_contract.V4PackageInputError
HarnessInput = input_contract.HarnessInput
HarnessOutput = input_contract.HarnessOutput
HarnessCaseOutput = input_contract.HarnessCaseOutput

_TRANSCRIPT_SCHEMA = "docwen.machine_session_transcript.v1"


def _frame(value: Mapping[str, object]) -> bytes:
    body = json.dumps(value, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    return f"Content-Length: {len(body)}\r\n\r\n".encode("ascii") + body


def _read_frame(stream: Any) -> dict[str, Any]:
    header = b""
    while b"\r\n\r\n" not in header:
        byte = stream.read(1)
        if not byte:
            raise V4PackageInputError("reverse_neutral_cli_closed_stdout")
        header += byte
    match = __import__("re").fullmatch(rb"Content-Length: ([1-9][0-9]*)\r\n\r\n", header)
    if match is None:
        raise V4PackageInputError("reverse_neutral_cli_frame_header_invalid")
    body = stream.read(int(match.group(1)))
    if len(body) != int(match.group(1)):
        raise V4PackageInputError("reverse_neutral_cli_frame_truncated")
    try:
        return json.loads(body.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise V4PackageInputError("reverse_neutral_cli_frame_json_invalid") from exc


def _inspect_docx(
    payload: bytes, expected: Mapping[str, int | bool], neutral: Mapping[str, object], plan: Mapping[str, object]
) -> dict[str, object]:
    try:
        return input_contract.inspect_docx(payload, expected, neutral, plan)
    except V4PackageInputError as exc:
        raise V4PackageInputError(f"reverse_neutral_headless_ooxml_failed:{exc}") from exc


def _default_harness_runner(
    executable: Path,
    docwen_clone: Path,
    run_root: Path,
    harness: HarnessInput,
) -> HarnessOutput:
    """Execute the frozen exact-two harness against the packaged DocWen CLI.

    The runner stays fail-closed when the packaged CLI is unavailable or the
    clone does not carry the resolved-v4 source, and it never substitutes
    legacy `convert.docx.to_markdown` output for the authenticated reverse
    recovery proof.  ``run_root`` must be a fresh directory owned by the
    producer.
    """

    if not executable.is_file() or executable.is_symlink():
        raise V4PackageInputError("reverse_neutral_cli_unavailable")
    if not (docwen_clone / "packages" / "core" / "src").is_dir():
        raise V4PackageInputError("reverse_neutral_docwen_source_unavailable")
    if run_root.exists() or run_root.is_symlink():
        raise V4PackageInputError("reverse_neutral_run_root_not_fresh")
    run_root.mkdir(parents=False)

    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    process = subprocess.Popen(
        [str(executable), "serve", "--stdio"],
        cwd=run_root,
        env=env,
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    if process.stdin is None or process.stdout is None or process.stderr is None:
        raise V4PackageInputError("reverse_neutral_stdio_unavailable")
    stdin = process.stdin
    stdout = process.stdout
    stderr = process.stderr
    events: list[dict[str, object]] = []
    response_frames: list[bytes] = []

    def request(operation: str, case_id: str | None, message: dict[str, object]) -> None:
        raw = _frame(message)
        stdin.write(raw)
        stdin.flush()
        events.append(
            machine_proof.transcript_event(
                direction="request",
                operation=operation,
                case_id=case_id,
                raw_frame=raw,
                message=message,
            )
        )

    def response(operation: str, case_id: str | None) -> dict[str, Any]:
        raw_message = _read_frame(stdout)
        raw = _frame(raw_message)
        response_frames.append(raw)
        events.append(
            machine_proof.transcript_event(
                direction="response",
                operation=operation,
                case_id=case_id,
                raw_frame=raw,
                message=raw_message,
            )
        )
        return raw_message

    try:
        request("initialize", None, {"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {}})
        initialize = response("initialize", None)
        if initialize.get("method") != "initialize":
            raise V4PackageInputError("reverse_neutral_initialize_invalid")
        request("discovery", None, {"jsonrpc": "2.0", "id": 2, "method": "capability/list", "params": {}})
        discovery = response("discovery", None)
        if discovery.get("method") != "capability/list":
            raise V4PackageInputError("reverse_neutral_discovery_invalid")
        discovery_result = discovery.get("result")
        capabilities = (
            discovery_result.get("capabilities")
            if isinstance(discovery_result, dict) and isinstance(discovery_result.get("capabilities"), list)
            else ()
        )
        capability_ids = {
            item.get("capability_id")
            for item in capabilities
            if isinstance(item, dict) and isinstance(item.get("capability_id"), str)
        }
        if input_contract.EXACT_CAPABILITY_ID not in capability_ids:
            raise V4PackageInputError("reverse_neutral_exact_capability_unadvertised")
        if "convert.docx.to_markdown" not in capability_ids:
            raise V4PackageInputError("reverse_neutral_reverse_capability_unadvertised")

        validation_plan_id = "plan.validation"
        request(
            "validation-plan",
            None,
            {
                "jsonrpc": "2.0",
                "id": 10,
                "method": "task/plan",
                "params": {
                    "capability_id": "validate.markdown",
                    "inputs": [{"role": "source", "locator": {"kind": "local_path", "path": str(docwen_clone)}}],
                    "output": {
                        "staging_root": {"kind": "local_path", "path": str(run_root)},
                        "staging_policy": "require_empty",
                    },
                    "options": {},
                },
            },
        )
        validation_plan = response("validation-plan", None)
        if not isinstance(validation_plan.get("result"), dict):
            raise V4PackageInputError("reverse_neutral_validation_plan_invalid")
        request(
            "validation-execute",
            None,
            {"jsonrpc": "2.0", "id": 11, "method": "task/execute", "params": {"plan_id": validation_plan_id}},
        )
        validation_execute = response("validation-execute", None)
        if validation_execute.get("method") != "task/execute":
            raise V4PackageInputError("reverse_neutral_validation_execute_invalid")
        validation_terminal = _read_terminal(stdout, "validation")
        if validation_terminal.get("method") != "task/completed":
            raise V4PackageInputError("reverse_neutral_validation_failed")
        events.append(
            machine_proof.transcript_event(
                direction="response",
                operation="validation-terminal",
                case_id=None,
                raw_frame=_frame(validation_terminal),
                message=machine_proof._normalize_message(validation_terminal),
            )
        )
        response_frames.append(_frame(validation_terminal))

        case_outputs: list[HarnessCaseOutput] = []
        for index, case in enumerate(harness.cases):
            request_base = 100 + index * 10
            suffix = case.case_id
            exact_inputs = [
                {
                    "input_id": "input.neutral-document",
                    "kind": "document",
                    "role": "neutral_document",
                    "logical_path": "inputs/document.resolved.json",
                    "locator": {"kind": "local_path", "path": str(case.neutral_path)},
                    "media_type": input_contract.NEUTRAL_MEDIA_TYPE,
                    "size_bytes": case.neutral_identity["bytes"],
                    "sha256": case.neutral_identity["sha256"],
                },
                {
                    "input_id": "input.numbering-export-plan",
                    "kind": "resource",
                    "role": "numbering_export_plan",
                    "logical_path": "inputs/numbering-export-plan.json",
                    "locator": {"kind": "local_path", "path": str(case.plan_path)},
                    "media_type": input_contract.PLAN_MEDIA_TYPE,
                    "size_bytes": case.plan_identity["bytes"],
                    "sha256": case.plan_identity["sha256"],
                },
            ]
            request(
                f"forward-plan:{suffix}",
                suffix,
                {
                    "jsonrpc": "2.0",
                    "id": request_base,
                    "method": "task/plan",
                    "params": {
                        "capability_id": input_contract.EXACT_CAPABILITY_ID,
                        "inputs": exact_inputs,
                        "output": {
                            "staging_root": {"kind": "local_path", "path": str(case.plan_path.parent)},
                            "staging_policy": "require_empty",
                        },
                        "options": {},
                    },
                },
            )
            forward_plan = response(f"forward-plan:{suffix}", suffix)
            if not isinstance(forward_plan.get("result"), dict):
                raise V4PackageInputError(f"reverse_neutral_forward_plan_invalid:{suffix}")
            request(
                f"forward-execute:{suffix}",
                suffix,
                {
                    "jsonrpc": "2.0",
                    "id": request_base + 1,
                    "method": "task/execute",
                    "params": {"plan_id": f"plan.forward.{suffix}"},
                },
            )
            forward_execute = response(f"forward-execute:{suffix}", suffix)
            if forward_execute.get("method") != "task/execute":
                raise V4PackageInputError(f"reverse_neutral_forward_execute_invalid:{suffix}")
            forward_terminal = _read_terminal(stdout, f"forward:{suffix}")
            if forward_terminal.get("method") != "task/completed":
                raise V4PackageInputError(f"reverse_neutral_forward_failed:{suffix}")
            events.append(
                machine_proof.transcript_event(
                    direction="response",
                    operation=f"forward-terminal:{suffix}",
                    case_id=suffix,
                    raw_frame=_frame(forward_terminal),
                    message=machine_proof._normalize_message(forward_terminal),
                )
            )
            response_frames.append(_frame(forward_terminal))

            forward_result = forward_terminal.get("params", {}).get("result")
            artifact_locators = forward_result.get("artifacts", []) if isinstance(forward_result, dict) else []
            docx_locator = next(
                (item for item in artifact_locators if isinstance(item, dict) and item.get("kind") == "document"),
                None,
            )
            if docx_locator is None or not isinstance(docx_locator.get("locator"), dict):
                raise V4PackageInputError(f"reverse_neutral_forward_artifact_missing:{suffix}")
            docx_path = Path(str(docx_locator["locator"].get("path", "")))
            if not docx_path.is_file():
                raise V4PackageInputError(f"reverse_neutral_forward_docx_missing:{suffix}")
            docx = docx_path.read_bytes()

            request(
                f"reverse-plan:{suffix}",
                suffix,
                {
                    "jsonrpc": "2.0",
                    "id": request_base + 2,
                    "method": "task/plan",
                    "params": {
                        "capability_id": "convert.docx.to_markdown",
                        "inputs": [
                            {
                                "input_id": "input.proof-docx",
                                "kind": "document",
                                "role": "source",
                                "logical_path": f"inputs/{suffix}.docx",
                                "locator": {"kind": "local_path", "path": str(docx_path)},
                                "media_type": input_contract.DOCX_MEDIA_TYPE,
                                "size_bytes": len(docx),
                                "sha256": hashlib.sha256(docx).hexdigest(),
                            }
                        ],
                        "output": {
                            "staging_root": {"kind": "local_path", "path": str(case.plan_path.parent)},
                            "staging_policy": "require_empty",
                        },
                        "options": {},
                    },
                },
            )
            reverse_plan = response(f"reverse-plan:{suffix}", suffix)
            if not isinstance(reverse_plan.get("result"), dict):
                raise V4PackageInputError(f"reverse_neutral_reverse_plan_invalid:{suffix}")
            request(
                f"reverse-execute:{suffix}",
                suffix,
                {
                    "jsonrpc": "2.0",
                    "id": request_base + 3,
                    "method": "task/execute",
                    "params": {"plan_id": f"plan.reverse.{suffix}"},
                },
            )
            reverse_execute = response(f"reverse-execute:{suffix}", suffix)
            if reverse_execute.get("method") != "task/execute":
                raise V4PackageInputError(f"reverse_neutral_reverse_execute_invalid:{suffix}")
            reverse_terminal = _read_terminal(stdout, f"reverse:{suffix}")
            if reverse_terminal.get("method") != "task/completed":
                raise V4PackageInputError(f"reverse_neutral_reverse_failed:{suffix}")
            events.append(
                machine_proof.transcript_event(
                    direction="response",
                    operation=f"reverse-terminal:{suffix}",
                    case_id=suffix,
                    raw_frame=_frame(reverse_terminal),
                    message=machine_proof._normalize_message(reverse_terminal),
                )
            )
            response_frames.append(_frame(reverse_terminal))

            reverse_result = reverse_terminal.get("params", {}).get("result")
            reverse_artifacts = reverse_result.get("artifacts", []) if isinstance(reverse_result, dict) else []
            md_locator = next(
                (item for item in reverse_artifacts if isinstance(item, dict) and item.get("kind") == "document"),
                None,
            )
            if md_locator is None or not isinstance(md_locator.get("locator"), dict):
                raise V4PackageInputError(f"reverse_neutral_reverse_artifact_missing:{suffix}")
            md_path = Path(str(md_locator["locator"].get("path", "")))
            if not md_path.is_file():
                raise V4PackageInputError(f"reverse_neutral_reverse_markdown_missing:{suffix}")
            roundtrip = md_path.read_bytes()

            inspection = _inspect_docx(docx, case.expected_ooxml, case.neutral_envelope, case.plan_envelope)
            case_outputs.append(
                HarnessCaseOutput(
                    case_id=case.case_id,
                    docx=docx,
                    roundtrip=roundtrip,
                    inspection=inspection,
                )
            )
        outputs = tuple(case_outputs)
        stdout_bytes = b"".join(response_frames)
        stderr_bytes = stderr.read()
        digest = machine_proof.transcript_request_digest(events)
        transcript = machine_proof.build_session_transcript(
            harness=harness,
            events=events,
            request_digest=digest,
            stdout=stdout_bytes,
            stderr=stderr_bytes,
            exit_code=process.returncode,
        )
        machine_proof.validate_session_transcript(transcript, harness=harness, outputs=outputs)
        return HarnessOutput(
            validation_terminal=validation_terminal,
            stdout=stdout_bytes,
            stderr=stderr_bytes,
            transcript=transcript,
            cases=outputs,
            request_digest=digest,
        )
    finally:
        stdin.close()
        try:
            process.wait(timeout=30)
        except subprocess.TimeoutExpired:
            process.kill()
            process.wait(timeout=10)


def _read_terminal(stream: Any, operation: str) -> dict[str, Any]:
    message = _read_frame(stream)
    if message.get("method") not in {"task/completed", "task/failed", "task/cancelled"}:
        raise V4PackageInputError(f"reverse_neutral_terminal_missing:{operation}")
    return message


__all__ = ["_default_harness_runner", "_frame"]
