"""Real Machine admission boundary for unresolved YAML-bearing Markdown."""

from __future__ import annotations

import hashlib
import io
from pathlib import Path

import pytest

from docwen_bundle.machine_factory import create_machine_server
from docwen_cli.machine.contracts import MachineContractValidator
from docwen_cli.machine.framing import FrameWriter, read_frame

pytestmark = pytest.mark.integration


def _request(method: str, request_id: int, params: dict[str, object]) -> dict[str, object]:
    return {"jsonrpc": "2.0", "id": request_id, "method": method, "params": params}


def test_machine_exact_two_plan_rejects_unresolved_yaml_source_without_leaking_content(
    tmp_path: Path,
) -> None:
    source_text = '\ufeff---\ntitle: "@yaml-citation [[Page#Heading]] ^yaml_bad"\n---\n\n前缀😀正文 ^bad_id\n'
    source_bytes = source_text.encode("utf-8")
    source_sha256 = hashlib.sha256(source_bytes).hexdigest()
    input_id = "input.yaml"
    source = tmp_path / "machine-yaml-invalid.md"
    source.write_bytes(source_bytes)
    staging = tmp_path / "staging"
    staging.mkdir()

    incoming = io.BytesIO()
    writer = FrameWriter(incoming)
    writer.write(
        _request(
            "initialize",
            1,
            {
                "protocol": {"name": "docwen.machine", "major": 1, "minor": 0},
                "client": {"name": "v4-yaml-gate", "version": "1.0.0"},
                "features": {"progress": True, "cancellation": True},
            },
        )
    )
    writer.write(
        _request(
            "task/plan",
            2,
            {
                "capability_id": "convert.markdown.to_docx",
                "inputs": [
                    {
                        "input_id": input_id,
                        "locator": {"kind": "local_path", "path": str(source)},
                        "kind": "document",
                        "role": "source",
                        "logical_path": "source.md",
                        "media_type": "text/markdown",
                        "size_bytes": len(source_bytes),
                        "sha256": source_sha256,
                    }
                ],
                "output": {
                    "staging_root": {"kind": "local_path", "path": str(staging)},
                    "staging_policy": "require_empty",
                },
                "options": {},
            },
        )
    )
    incoming.seek(0)
    outgoing = io.BytesIO()

    server = create_machine_server(reader=incoming, writer=outgoing)
    assert server.run() == 0

    outgoing.seek(0)
    messages: list[dict[str, object]] = []
    while (message := read_frame(outgoing)) is not None:
        MachineContractValidator().validate_message(message)
        messages.append(message)
    assert len(messages) == 2, messages
    rejected = messages[1]
    assert rejected["error"]["code"] == -32602  # type: ignore[index]
    task_error = rejected["error"]["data"]["task_error"]  # type: ignore[index]
    assert (task_error["category"], task_error["code"], task_error["retryable"]) == (
        "invalid_request",
        "undeclared_input_role",
        False,
    )
    assert source_text not in repr(rejected)
    assert str(source) not in repr(rejected)
    assert list(staging.iterdir()) == []
