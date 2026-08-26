from __future__ import annotations

import hashlib

from scripts.release import build_v4_package_input as producer


def synthetic_transcript(
    harness: producer.HarnessInput,
    outputs: tuple[producer.HarnessCaseOutput, ...],
    validation_terminal: dict[str, object],
) -> tuple[bytes, bytes, str]:
    events: list[dict[str, object]] = []
    response_frames: list[bytes] = []

    def request(operation: str, case_id: str | None, message: dict[str, object]) -> None:
        raw = producer._frame(message)
        events.append(
            producer._transcript_event(
                direction="request",
                operation=operation,
                case_id=case_id,
                raw_frame=raw,
                message=message,
            )
        )

    def response(operation: str, case_id: str | None, message: dict[str, object]) -> None:
        raw = producer._frame(message)
        response_frames.append(raw)
        events.append(
            producer._transcript_event(
                direction="response",
                operation=operation,
                case_id=case_id,
                raw_frame=raw,
                message=message,
            )
        )

    request("initialize", None, {"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {}})
    response("initialize", None, {"jsonrpc": "2.0", "id": 1, "result": {}})
    request("discovery", None, {"jsonrpc": "2.0", "id": 2, "method": "capability/list", "params": {}})
    response("discovery", None, {"jsonrpc": "2.0", "id": 2, "result": {}})
    validation_plan = "plan.validation"
    validation_task = str(validation_terminal["params"]["task_id"])
    request(
        "validation-plan",
        None,
        {
            "jsonrpc": "2.0",
            "id": 10,
            "method": "task/plan",
            "params": {
                "capability_id": "validate.markdown",
                "inputs": [{"role": "source"}],
                "output": {
                    "staging_root": {"kind": "local_path", "path": "C:/synthetic/validation"},
                    "staging_policy": "require_empty",
                },
                "options": {},
            },
        },
    )
    response("validation-plan", None, {"jsonrpc": "2.0", "id": 10, "result": {"plan_id": validation_plan}})
    request(
        "validation-execute",
        None,
        {"jsonrpc": "2.0", "id": 11, "method": "task/execute", "params": {"plan_id": validation_plan}},
    )
    response(
        "validation-execute",
        None,
        {"jsonrpc": "2.0", "id": 11, "result": {"task_id": validation_task, "state": "accepted"}},
    )
    response("validation-terminal", None, validation_terminal)
    output_by_case = {item.case_id: item for item in outputs}
    for index, case in enumerate(harness.cases):
        suffix = case.case_id
        request_base = 100 + (index * 10)
        forward_plan = f"plan.forward.{suffix}"
        forward_task = f"task.forward.{suffix}"
        reverse_plan = f"plan.reverse.{suffix}"
        reverse_task = f"task.reverse.{suffix}"
        exact_inputs = [
            {
                "input_id": "input.neutral-document",
                "kind": "document",
                "role": "neutral_document",
                "logical_path": "inputs/document.resolved.json",
                "locator": {"kind": "local_path", "path": str(case.neutral_path)},
                "media_type": producer.NEUTRAL_MEDIA_TYPE,
                "size_bytes": case.neutral_identity["bytes"],
                "sha256": case.neutral_identity["sha256"],
            },
            {
                "input_id": "input.numbering-export-plan",
                "kind": "resource",
                "role": "numbering_export_plan",
                "logical_path": "inputs/numbering-export-plan.json",
                "locator": {"kind": "local_path", "path": str(case.plan_path)},
                "media_type": producer.PLAN_MEDIA_TYPE,
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
                    "capability_id": producer.EXACT_CAPABILITY_ID,
                    "inputs": exact_inputs,
                    "output": {
                        "staging_root": {"kind": "local_path", "path": str(case.plan_path.parent)},
                        "staging_policy": "require_empty",
                    },
                    "options": {},
                },
            },
        )
        response(
            f"forward-plan:{suffix}",
            suffix,
            {"jsonrpc": "2.0", "id": request_base, "result": {"plan_id": forward_plan}},
        )
        request(
            f"forward-execute:{suffix}",
            suffix,
            {
                "jsonrpc": "2.0",
                "id": request_base + 1,
                "method": "task/execute",
                "params": {"plan_id": forward_plan},
            },
        )
        response(
            f"forward-execute:{suffix}",
            suffix,
            {
                "jsonrpc": "2.0",
                "id": request_base + 1,
                "result": {"task_id": forward_task, "state": "accepted"},
            },
        )
        response(
            f"forward-terminal:{suffix}",
            suffix,
            {"jsonrpc": "2.0", "method": "task/completed", "params": {"task_id": forward_task}},
        )
        output = output_by_case[suffix]
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
                            "locator": {"kind": "local_path", "path": str(case.plan_path)},
                            "media_type": producer.DOCX_MEDIA_TYPE,
                            "size_bytes": len(output.docx),
                            "sha256": hashlib.sha256(output.docx).hexdigest(),
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
        response(
            f"reverse-plan:{suffix}",
            suffix,
            {"jsonrpc": "2.0", "id": request_base + 2, "result": {"plan_id": reverse_plan}},
        )
        request(
            f"reverse-execute:{suffix}",
            suffix,
            {
                "jsonrpc": "2.0",
                "id": request_base + 3,
                "method": "task/execute",
                "params": {"plan_id": reverse_plan},
            },
        )
        response(
            f"reverse-execute:{suffix}",
            suffix,
            {
                "jsonrpc": "2.0",
                "id": request_base + 3,
                "result": {"task_id": reverse_task, "state": "accepted"},
            },
        )
        response(
            f"reverse-terminal:{suffix}",
            suffix,
            {"jsonrpc": "2.0", "method": "task/completed", "params": {"task_id": reverse_task}},
        )
    stdout = b"".join(response_frames)
    digest = producer._transcript_request_digest(events)
    transcript = producer._build_session_transcript(
        harness=harness,
        events=events,
        request_digest=digest,
        stdout=stdout,
        stderr=b"",
        exit_code=0,
    )
    return transcript, stdout, digest


__all__ = ["synthetic_transcript"]
