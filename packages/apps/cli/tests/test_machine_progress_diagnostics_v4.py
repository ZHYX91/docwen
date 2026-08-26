"""Machine v1 bounded progress and v4 diagnostic evidence contracts."""

from __future__ import annotations

import io
import json
from collections.abc import Callable, Mapping
from dataclasses import dataclass
from typing import Any

import pytest

from docwen_application.conversion_service import ConversionTaskOutcome
from docwen_cli.machine.contracts import MachineContractValidator
from docwen_cli.machine.framing import FrameWriter, read_frame
from docwen_cli.machine.server import MachineProtocolServer
from docwen_core.models import (
    ArtifactBundle,
    BundleArtifact,
    BundleEntry,
    BundleProducer,
    ConversionDiagnostic,
    ConversionErrorInfo,
    ConversionMetrics,
    DiagnosticFix,
    DiagnosticRange,
    DiagnosticSource,
    DiagnosticTextEdit,
    TaskEvent,
)

pytestmark = pytest.mark.contract


@dataclass(frozen=True)
class _Plan:
    payload: dict[str, Any]

    def to_dict(self) -> dict[str, Any]:
        return dict(self.payload)


class _Service:
    def __init__(
        self,
        *,
        events: tuple[TaskEvent, ...] = (),
        outcome_factory: Callable[[str], ConversionTaskOutcome] | None = None,
    ) -> None:
        self.events = events
        self.outcome_factory = outcome_factory or _completed_outcome
        self.event_callback: Callable[[TaskEvent], None] | None = None

    def plan(self, request: Any) -> _Plan:
        return _Plan(
            {
                "plan_id": "plan.test",
                "capability_id": request.capability_id,
                "effective_options": {},
                "output_shape": {
                    "cardinality": "one",
                    "artifact_kinds": ["document"],
                    "relation_types": [],
                    "atomic_bundle": True,
                },
                "warnings": [],
                "limitations": [],
                "requires_confirmation": False,
            }
        )

    def accept(self, plan_id: str) -> str:
        assert plan_id == "plan.test"
        return "task.test"

    def execute_accepted(self, task_id: str) -> ConversionTaskOutcome:
        assert self.event_callback is not None
        for event in self.events:
            self.event_callback(event)
        return self.outcome_factory(task_id)


def _completed_outcome(task_id: str) -> ConversionTaskOutcome:
    artifact = BundleArtifact(
        artifact_id="artifact.primary",
        kind="document",
        locator="result.docx",
        suggested_name="result.docx",
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        size_bytes=4,
        sha256="1" * 64,
    )
    bundle = ArtifactBundle(
        bundle_id="bundle.test",
        task_id=task_id,
        producer=BundleProducer(product_version="0.9.0"),
        artifacts=(artifact,),
        entries=(BundleEntry(artifact.artifact_id, "primary", 0, True),),
    )
    return ConversionTaskOutcome(
        task_id=task_id,
        state="completed",
        bundle=bundle,
        diagnostics=(),
        metrics=ConversionMetrics(duration_ms=1, input_bytes=3, output_bytes=4),
    )


def _request(method: str, request_id: int, params: dict[str, Any]) -> dict[str, Any]:
    return {"jsonrpc": "2.0", "id": request_id, "method": method, "params": params}


def _messages(capability_id: str = "convert.markdown.to_docx") -> list[dict[str, Any]]:
    return [
        _request(
            "initialize",
            1,
            {
                "protocol": {"name": "docwen.machine", "major": 1, "minor": 0},
                "client": {"name": "test-client", "version": "1.0.0"},
                "features": {"progress": True, "cancellation": True},
            },
        ),
        _request(
            "task/plan",
            2,
            {
                "capability_id": capability_id,
                "inputs": [
                    {
                        "input_id": "input.1",
                        "locator": {"kind": "local_path", "path": r"C:\fixture\source.md"},
                        "kind": "document",
                        "role": "source",
                        "logical_path": "source.md",
                        "media_type": "text/markdown",
                        "size_bytes": 3,
                        "sha256": "0" * 64,
                    }
                ],
                "output": {
                    "staging_root": {"kind": "local_path", "path": r"C:\fixture\staging"},
                    "staging_policy": "require_empty",
                },
                "options": {},
            },
        ),
        _request("task/execute", 3, {"plan_id": "plan.test"}),
    ]


def _run(
    service: _Service,
    *,
    capability_id: str = "convert.markdown.to_docx",
) -> tuple[list[dict[str, Any]], Mapping[str, object]]:
    incoming = io.BytesIO()
    frame_writer = FrameWriter(incoming)
    for message in _messages(capability_id):
        frame_writer.write(message)
    incoming.seek(0)
    outgoing = io.BytesIO()
    server = MachineProtocolServer(service, incoming, outgoing)  # type: ignore[arg-type]
    service.event_callback = server.report_runtime_event

    assert server.run() == 0
    active_state = dict(server._task_progress)
    outgoing.seek(0)
    responses: list[dict[str, Any]] = []
    while (message := read_frame(outgoing)) is not None:
        responses.append(message)
    return responses, active_state


def _lifecycle(responses: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [
        message
        for message in responses
        if message.get("method") in {"task/progress", "task/completed", "task/failed", "task/cancelled"}
    ]


def test_runtime_progress_is_private_bounded_monotonic_and_terminal_ordered() -> None:
    private_path = r"C:\private\customers\source.md"
    private_text = "confidential authored paragraph"
    events = (
        TaskEvent("task.test", "task_progress", 1, payload={"percent": 10, "message": private_path}),
        TaskEvent("task.test", "task_progress", 2, payload={"percent": 10, "message": private_text}),
        TaskEvent("task.test", "task_progress", 3, payload={"percent": 5, "message": private_text}),
        TaskEvent("task.test", "task_progress", 4, payload={"percent": True, "message": private_text}),
        TaskEvent("task.test", "task_progress", 5, payload={"percent": float("nan")}),
        TaskEvent("task.test", "task_progress", 6, payload={"percent": 50}),
        TaskEvent("task.test", "task_progress", 7, payload={"percent": 100, "message": private_path}),
        TaskEvent("task.other", "task_progress", 8, payload={"percent": 80}),
        TaskEvent("task.test", "diagnostic", 9, payload={"message": private_text, "location": private_path}),
    )

    responses, active = _run(_Service(events=events))

    lifecycle = _lifecycle(responses)
    assert [message["method"] for message in lifecycle] == ["task/progress"] * 5 + ["task/completed"]
    assert [message["params"]["completed"] for message in lifecycle[:-1]] == [0, 10, 50, 95, 100]
    assert [message["params"]["sequence"] for message in lifecycle] == [1, 2, 3, 4, 5, 6]
    assert {message["params"]["phase"] for message in lifecycle[:-1]} == {"conversion"}
    assert private_path not in json.dumps(responses, ensure_ascii=False)
    assert private_text not in json.dumps(responses, ensure_ascii=False)
    assert active == {}
    validator = MachineContractValidator()
    for response in responses:
        validator.validate_message(response)


def test_validation_uses_closed_progress_phase_and_messages() -> None:
    events = (
        TaskEvent("task.test", "task_progress", 1, payload={"percent": 25, "message": "secret.md"}),
        TaskEvent("task.test", "task_progress", 2, payload={"percent": 75, "message": "private text"}),
    )

    responses, _ = _run(_Service(events=events), capability_id="validate.markdown")

    progress = [message["params"] for message in responses if message.get("method") == "task/progress"]
    assert [item["completed"] for item in progress] == [0, 25, 75, 100]
    assert [item["message"] for item in progress] == [
        "Validation started",
        "Validation progress 25 percent",
        "Validation progress 75 percent",
        "Validation complete",
    ]
    assert {item["phase"] for item in progress} == {"conversion"}


def _failed_semantic_outcome(task_id: str) -> ConversionTaskOutcome:
    return ConversionTaskOutcome(
        task_id=task_id,
        state="failed",
        bundle=None,
        diagnostics=(
            ConversionDiagnostic(
                level="error",
                message="Standalone block anchor has no preceding structured block.",
                code="docwen.markdown.anchor.dangling",
                evidence_schema="docwen.machine.diagnostic_evidence.v1",
                source=DiagnosticSource(input_id="input.1", sha256="0" * 64),
                range=DiagnosticRange(6, 14),
                related_ranges=(),
                fixes=(),
            ),
        ),
        metrics=ConversionMetrics(input_bytes=14),
        error=ConversionErrorInfo(
            error_type="invalid_request",
            message="Markdown semantics are invalid.",
            diagnostic_code="invalid_document_semantics",
        ),
    )


def test_failed_semantic_diagnostic_carries_complete_source_evidence() -> None:
    responses, active = _run(_Service(outcome_factory=_failed_semantic_outcome))

    failed = next(message for message in responses if message.get("method") == "task/failed")
    assert failed["params"]["diagnostics"] == [
        {
            "severity": "error",
            "code": "docwen.markdown.anchor.dangling",
            "message": "Standalone block anchor has no preceding structured block.",
            "evidence_schema": "docwen.machine.diagnostic_evidence.v1",
            "source": {
                "input_id": "input.1",
                "sha256": "0" * 64,
                "encoding": "utf-8",
                "coordinate_system": "unicode_code_point",
                "offset_base": 0,
                "range_end": "exclusive",
            },
            "range": {"start": 6, "end": 14},
            "related_ranges": [],
            "fixes": [],
        }
    ]
    assert active == {}
    MachineContractValidator().validate_message(failed)


def test_fix_evidence_round_trips_without_reordering_edits() -> None:
    diagnostic = ConversionDiagnostic(
        level="error",
        message="Invalid identifier.",
        code="docwen.markdown.anchor.invalid_id",
        evidence_schema="docwen.machine.diagnostic_evidence.v1",
        source=DiagnosticSource(input_id="input.1", sha256="0" * 64),
        range=DiagnosticRange(4, 13),
        fixes=(
            DiagnosticFix(
                fix_id="docwen.markdown.fix.replace_invalid_id",
                edits=(
                    DiagnosticTextEdit(DiagnosticRange(4, 13), "valid-id"),
                    DiagnosticTextEdit(DiagnosticRange(20, 20), " ^valid-id"),
                ),
            ),
        ),
    )

    restored = ConversionDiagnostic.from_dict(diagnostic.to_dict())

    assert restored.to_dict() == diagnostic.to_dict()
    assert [edit.range.start for edit in restored.fixes[0].edits] == [4, 20]


@pytest.mark.parametrize("invalid_kind", ["missing_evidence", "wrong_source"])
def test_invalid_semantic_outcome_fails_terminally_and_releases_admission(invalid_kind: str) -> None:
    def outcome(task_id: str) -> ConversionTaskOutcome:
        source = None
        evidence_schema = None
        if invalid_kind == "wrong_source":
            source = DiagnosticSource(input_id="input.other", sha256="2" * 64)
            evidence_schema = "docwen.machine.diagnostic_evidence.v1"
        return ConversionTaskOutcome(
            task_id=task_id,
            state="completed",
            bundle=_completed_outcome(task_id).bundle,
            diagnostics=(
                ConversionDiagnostic(
                    level="error",
                    message="Invalid semantic evidence.",
                    code="docwen.markdown.anchor.dangling",
                    evidence_schema=evidence_schema,
                    source=source,
                    range=DiagnosticRange(0, 3) if source is not None else None,
                ),
            ),
            metrics=ConversionMetrics(),
        )

    responses, active = _run(_Service(outcome_factory=outcome))

    lifecycle = _lifecycle(responses)
    assert [message["method"] for message in lifecycle] == ["task/progress", "task/failed"]
    assert lifecycle[0]["params"]["completed"] == 0
    assert lifecycle[1]["params"]["error"]["code"] == "internal_error"
    assert lifecycle[1]["params"]["diagnostics"] == []
    assert active == {}
