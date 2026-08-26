"""Machine Protocol v1 server lifecycle and schema conformance contracts."""

from __future__ import annotations

import io
import threading
from dataclasses import dataclass
from typing import Any

import pytest
from tools.validate_contracts import validate_trace

from docwen_application.conversion_service import ConversionTaskOutcome
from docwen_cli.machine.contracts import MachineContractValidator
from docwen_cli.machine.framing import MAX_MESSAGE_BYTES, FrameWriter, read_frame
from docwen_cli.machine.server import MachineProtocolServer
from docwen_core.models import (
    ArtifactBundle,
    BundleArtifact,
    BundleEntry,
    BundleProducer,
    ConversionDiagnostic,
    ConversionErrorInfo,
    ConversionMetrics,
)

pytestmark = pytest.mark.contract


@dataclass(frozen=True)
class _WireObject:
    payload: dict[str, Any]

    def to_dict(self) -> dict[str, Any]:
        return dict(self.payload)


class _Service:
    def __init__(self, *, cancelled: bool = False) -> None:
        self.cancelled = cancelled
        self.cancel_called = threading.Event()

    def list_capabilities(self) -> tuple[_WireObject, ...]:
        return (
            _WireObject(
                {
                    "capability_id": "convert.markdown.to_docx",
                    "operation": "convert",
                    "input_shape": {
                        "slots": [
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
                        ],
                        "undeclared_roles": "reject",
                    },
                    "output_media_types": ["application/vnd.openxmlformats-officedocument.wordprocessingml.document"],
                    "output_shape": {
                        "cardinality": "one",
                        "artifact_kinds": ["document"],
                        "relation_types": [],
                        "atomic_bundle": True,
                    },
                    "options_schema": {"type": "object", "additionalProperties": False},
                    "availability": "available",
                    "dependencies": [],
                    "limitations": [],
                }
            ),
        )

    def plan(self, request: Any) -> _WireObject:
        return _WireObject(
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
        if self.cancelled:
            self.cancel_called.wait(timeout=2)
            return ConversionTaskOutcome(
                task_id=task_id,
                state="cancelled",
                bundle=None,
                diagnostics=(),
                metrics=ConversionMetrics(),
                error=ConversionErrorInfo(error_type="cancelled", message="user_cancelled"),
            )
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

    def cancel(self, task_id: str) -> str:
        assert task_id == "task.test"
        self.cancel_called.set()
        return "cancellation_requested"


class _QueryService:
    def health_check(self) -> dict[str, Any]:
        return {
            "all_ok": True,
            "checks": [{"id": "config.load", "kind": "config", "label": "Config", "status": "ok", "reason": None}],
        }

    def inspect_file(self, input_handle: dict[str, Any]) -> dict[str, Any]:
        return {
            "file_path": input_handle["locator"]["path"],
            "size_bytes": input_handle["size_bytes"],
            "content_sha256": input_handle["sha256"],
            "decision": "allow",
            "supported_actions": ["inspect", "convert"],
            "declared_format": "md",
            "detected_format": "md",
            "warning_code": "",
            "reason_code": "",
            "workflow_category": "markdown",
        }

    def list_resources(self, kind: str, *, target: str | None, locale: str | None) -> dict[str, Any]:
        assert target is None
        assert locale == "en_US"
        return {
            "kind": kind,
            "resources": [{"id": "standard", "name": "Standard", "description": "Default"}],
        }

    def gui_control(
        self,
        action: str,
        *,
        file_path: str | None,
        timeout_seconds: int,
    ) -> dict[str, Any]:
        assert file_path is None
        assert timeout_seconds == 5
        return {"action": action, "accepted": True, "running": False}


def _request(method: str, request_id: int, params: dict[str, Any]) -> dict[str, Any]:
    return {"jsonrpc": "2.0", "id": request_id, "method": method, "params": params}


def _initialize() -> dict[str, Any]:
    return _request(
        "initialize",
        1,
        {
            "protocol": {"name": "docwen.machine", "major": 1, "minor": 0},
            "client": {"name": "test-client", "version": "1.0.0"},
            "features": {"progress": True, "cancellation": True},
        },
    )


def _plan() -> dict[str, Any]:
    return _request(
        "task/plan",
        3,
        {
            "capability_id": "convert.markdown.to_docx",
            "inputs": [
                {
                    "input_id": "input.neutral",
                    "locator": {
                        "kind": "local_path",
                        "path": "C:\\fixture\\document.resolved.json",
                    },
                    "kind": "document",
                    "role": "neutral_document",
                    "logical_path": "document.resolved.json",
                    "media_type": "application/vnd.docwen.resolved-document+json",
                    "size_bytes": 3,
                    "sha256": "0" * 64,
                },
                {
                    "input_id": "input.plan",
                    "locator": {
                        "kind": "local_path",
                        "path": "C:\\fixture\\numbering-plan.json",
                    },
                    "kind": "resource",
                    "role": "numbering_export_plan",
                    "logical_path": "numbering-plan.json",
                    "media_type": "application/vnd.docwen.numbering-export-plan+json",
                    "size_bytes": 3,
                    "sha256": "1" * 64,
                },
            ],
            "output": {
                "staging_root": {"kind": "local_path", "path": "C:\\fixture\\staging"},
                "staging_policy": "require_empty",
            },
            "options": {},
        },
    )


def _run(
    messages: list[dict[str, Any]],
    service: _Service,
    query_service: Any | None = None,
) -> tuple[int, list[dict[str, Any]]]:
    incoming = io.BytesIO()
    frame_writer = FrameWriter(incoming)
    for message in messages:
        frame_writer.write(message)
    incoming.seek(0)
    outgoing = io.BytesIO()

    server = MachineProtocolServer(service, incoming, outgoing, query_service=query_service)  # type: ignore[arg-type]
    exit_code = server.run()
    outgoing.seek(0)
    responses: list[dict[str, Any]] = []
    while (message := read_frame(outgoing)) is not None:
        responses.append(message)
    return exit_code, responses


def test_initialize_discovery_plan_execute_and_terminal_trace() -> None:
    requests = [
        _initialize(),
        _request("capability/list", 2, {}),
        _plan(),
        _request("task/execute", 4, {"plan_id": "plan.test"}),
    ]
    exit_code, responses = _run(requests, _Service())

    assert exit_code == 0
    assert responses[0]["result"]["artifact_bundle_schema"] == "docwen.artifact_bundle.v2"
    assert responses[1]["result"]["capabilities"][0]["capability_id"] == "convert.markdown.to_docx"
    assert responses[3]["result"] == {"task_id": "task.test", "state": "accepted"}
    assert responses[-1]["method"] == "task/completed"
    validator = MachineContractValidator()
    for message in responses:
        validator.validate_message(message)
    validate_trace([*requests, *responses], requires_terminal=True)


def test_plan_accepts_exact_resolved_numbering_handles() -> None:
    plan = _plan()

    exit_code, responses = _run([_initialize(), plan], _Service())

    assert exit_code == 0
    assert responses[-1]["result"]["plan_id"] == "plan.test"


def test_plan_schema_rejects_unknown_kind_role_and_invalid_logical_path() -> None:
    for field, value in (
        ("kind", "unknown"),
        ("role", "unknown_role"),
        ("logical_path", "document/../source.md"),
    ):
        plan = _plan()
        plan["params"]["inputs"][0][field] = value

        exit_code, responses = _run([_initialize(), plan], _Service())

        assert exit_code == 0
        assert responses[-1]["error"]["code"] == -32602
        assert responses[-1]["error"]["message"] == "Invalid params"


def test_cancel_request_produces_exactly_one_cancelled_terminal() -> None:
    requests = [
        _initialize(),
        _plan(),
        _request("task/execute", 4, {"plan_id": "plan.test"}),
        _request("task/cancel", 5, {"task_id": "task.test"}),
    ]
    exit_code, responses = _run(requests, _Service(cancelled=True))

    assert exit_code == 0
    assert any(message.get("result", {}).get("state") == "cancellation_requested" for message in responses)
    terminal = [
        message for message in responses if message.get("method") in {"task/completed", "task/failed", "task/cancelled"}
    ]
    assert [message["method"] for message in terminal] == ["task/cancelled"]


def test_completed_diagnostic_projects_artifact_binding() -> None:
    class _DiagnosticService(_Service):
        def execute_accepted(self, task_id: str) -> ConversionTaskOutcome:
            outcome = super().execute_accepted(task_id)
            return ConversionTaskOutcome(
                task_id=outcome.task_id,
                state=outcome.state,
                bundle=outcome.bundle,
                diagnostics=(
                    ConversionDiagnostic(
                        level="warning",
                        message="Page ownership could not be proven.",
                        code="resource_page_unresolved",
                        artifact_id="artifact.primary",
                    ),
                ),
                metrics=outcome.metrics,
                error=outcome.error,
            )

    requests = [_initialize(), _plan(), _request("task/execute", 4, {"plan_id": "plan.test"})]
    exit_code, responses = _run(requests, _DiagnosticService())

    assert exit_code == 0
    completed = next(message for message in responses if message.get("method") == "task/completed")
    assert completed["params"]["diagnostics"][0]["artifact_id"] == "artifact.primary"


def test_typed_query_and_gui_methods_are_schema_valid() -> None:
    input_handle = {
        "input_id": "input.query",
        "locator": {"kind": "local_path", "path": "C:\\fixture\\source.md"},
        "kind": "document",
        "role": "source",
        "logical_path": "source.md",
        "media_type": "text/markdown",
        "size_bytes": 3,
        "sha256": "0" * 64,
    }
    requests = [
        _initialize(),
        _request("health/check", 2, {}),
        _request("file/inspect", 3, {"input": input_handle}),
        _request("resource/list", 4, {"kind": "numbering-schemes", "locale": "en_US"}),
        _request("gui/status", 5, {"timeout_seconds": 5}),
    ]
    exit_code, responses = _run(requests, _Service(), _QueryService())

    assert exit_code == 0
    assert responses[1]["result"]["all_ok"] is True
    assert responses[2]["result"]["size_bytes"] == 3
    assert responses[2]["result"]["content_sha256"] == "0" * 64
    assert responses[3]["result"]["resources"][0]["id"] == "standard"
    assert responses[4]["result"] == {"action": "status", "accepted": True, "running": False}


def test_method_before_initialize_is_rejected_without_starting_task() -> None:
    request = _request("capability/list", 1, {})
    _, responses = _run([request], _Service())

    assert responses == [
        {
            "jsonrpc": "2.0",
            "id": 1,
            "error": {"code": -32600, "message": "initialize must be called first"},
        }
    ]


def test_malformed_frame_returns_parse_error_and_nonzero_exit() -> None:
    incoming = io.BytesIO(b"Content-Length: 3\r\n\r\n{}")
    outgoing = io.BytesIO()
    server = MachineProtocolServer(_Service(), incoming, outgoing)  # type: ignore[arg-type]

    assert server.run() == 2
    outgoing.seek(0)
    response = read_frame(outgoing)
    assert response is not None
    assert response["error"]["code"] == -32700
    assert response["id"] is None


def test_oversized_frame_returns_parse_error_without_service_action() -> None:
    class NoActionService:
        called = False

        def __getattr__(self, name: str) -> Any:
            self.called = True
            raise AssertionError(f"service action must not run for oversized frame: {name}")

    incoming = io.BytesIO(f"Content-Length: {MAX_MESSAGE_BYTES + 1}\r\n\r\n".encode("ascii"))
    outgoing = io.BytesIO()
    service = NoActionService()
    server = MachineProtocolServer(service, incoming, outgoing)  # type: ignore[arg-type]

    assert server.run() == 2
    assert service.called is False
    outgoing.seek(0)
    response = read_frame(outgoing)
    assert response == {
        "jsonrpc": "2.0",
        "id": None,
        "error": {"code": -32700, "message": "Parse error", "data": {"code": "frame_too_large"}},
    }
