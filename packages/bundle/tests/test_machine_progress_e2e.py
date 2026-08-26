"""Real Runtime-to-Machine progress projection for convert and validate tasks."""

from __future__ import annotations

import hashlib
import json
import re
import subprocess
import sys
from pathlib import Path
from typing import Any, BinaryIO, cast

import pytest
from scripts.release.verify_packaged_cli import (
    MACHINE_EXACT_TWO_NEUTRAL_DOCUMENT,
    MACHINE_EXACT_TWO_NUMBERING_PLAN,
    MACHINE_NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
    MACHINE_RESOLVED_DOCUMENT_MEDIA_TYPE,
)
from tools.validate_contracts import validate_trace

from docwen_cli.machine.contracts import MachineContractValidator
from docwen_cli.machine.framing import FrameWriter, read_frame

pytestmark = pytest.mark.e2e

_TERMINAL_METHODS = {"task/completed", "task/failed", "task/cancelled"}


def _request(request_id: int, method: str, params: dict[str, Any]) -> dict[str, Any]:
    return {"jsonrpc": "2.0", "id": request_id, "method": method, "params": params}


def test_real_stdio_convert_and_validate_emit_safe_advancing_progress(tmp_path: Path) -> None:
    process = subprocess.Popen(
        [
            sys.executable,
            "-c",
            "from docwen_bundle.cli_entry import main; raise SystemExit(main(['serve', '--stdio']))",
        ],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    assert process.stdin is not None
    assert process.stdout is not None
    assert process.stderr is not None
    process_stdin = cast(BinaryIO, process.stdin)
    process_stdout = cast(BinaryIO, process.stdout)
    writer = FrameWriter(process_stdin)
    trace: list[dict[str, Any]] = []
    next_request_id = 1
    stdin_closed = False

    def exchange(method: str, params: dict[str, Any]) -> dict[str, Any]:
        nonlocal next_request_id
        request = _request(next_request_id, method, params)
        next_request_id += 1
        trace.append(request)
        writer.write(request)
        response = read_frame(process_stdout)
        assert response is not None
        trace.append(response)
        return response

    def execute(
        capability_id: str,
        inputs: list[dict[str, Any]],
        staging: Path,
        *,
        options: dict[str, Any],
        label: str,
    ) -> list[dict[str, Any]]:
        planned = exchange(
            "task/plan",
            {
                "capability_id": capability_id,
                "inputs": inputs,
                "output": {
                    "staging_root": {"kind": "local_path", "path": str(staging)},
                    "staging_policy": "require_empty",
                },
                "options": options,
            },
        )
        accepted = exchange("task/execute", {"plan_id": planned["result"]["plan_id"]})
        task_id = accepted["result"]["task_id"]
        lifecycle: list[dict[str, Any]] = []
        while True:
            notification = read_frame(process_stdout)
            assert notification is not None
            trace.append(notification)
            lifecycle.append(notification)
            if notification.get("method") in _TERMINAL_METHODS:
                break

        progress = [message for message in lifecycle if message.get("method") == "task/progress"]
        terminal = lifecycle[-1]
        completed = [message["params"]["completed"] for message in progress]
        sequences = [message["params"]["sequence"] for message in lifecycle]

        assert accepted["result"]["state"] == "accepted"
        assert terminal["method"] == "task/completed", terminal
        assert completed[0] == 0
        assert completed[-1] == 100
        assert any(0 < value < 100 for value in completed)
        assert completed == sorted(set(completed))
        assert sequences == list(range(1, len(lifecycle) + 1))
        assert {message["params"]["task_id"] for message in lifecycle} == {task_id}
        assert {message["params"]["phase"] for message in progress} == {"conversion"}
        assert {message["params"]["total"] for message in progress} == {100}
        assert {message["params"]["unit"] for message in progress} == {"percent"}
        for message in progress:
            text = message["params"]["message"]
            assert re.fullmatch(rf"{label} (?:started|complete|progress [1-9][0-9]? percent)", text)
            assert len(text) <= 31
            assert len(json.dumps(message, separators=(",", ":")).encode("utf-8")) <= 384
            for item in inputs:
                input_path = Path(item["locator"]["path"])
                assert str(input_path) not in repr(message)
                assert input_path.read_text(encoding="utf-8") not in repr(message)
        return lifecycle

    try:
        initialized = exchange(
            "initialize",
            {
                "protocol": {"name": "docwen.machine", "major": 1, "minor": 0},
                "client": {"name": "docwen-progress-e2e", "version": "1.0.0"},
                "features": {"progress": True, "cancellation": True},
            },
        )
        assert initialized["result"]["features"]["progress"] is True

        neutral_document = tmp_path / "private-resolved-document.json"
        neutral_document.write_text(
            json.dumps(MACHINE_EXACT_TWO_NEUTRAL_DOCUMENT, ensure_ascii=False, separators=(",", ":")),
            encoding="utf-8",
        )
        numbering_plan = tmp_path / "private-numbering-export-plan.json"
        numbering_plan.write_text(
            json.dumps(MACHINE_EXACT_TWO_NUMBERING_PLAN, ensure_ascii=False, separators=(",", ":")),
            encoding="utf-8",
        )
        neutral_bytes = neutral_document.read_bytes()
        numbering_bytes = numbering_plan.read_bytes()
        convert_staging = tmp_path / "convert-staging"
        convert_staging.mkdir()
        convert_lifecycle = execute(
            "convert.markdown.to_docx",
            [
                {
                    "input_id": "input.conversion.neutral-document",
                    "locator": {"kind": "local_path", "path": str(neutral_document)},
                    "kind": "document",
                    "role": "neutral_document",
                    "logical_path": "inputs/document.resolved.json",
                    "media_type": MACHINE_RESOLVED_DOCUMENT_MEDIA_TYPE,
                    "size_bytes": len(neutral_bytes),
                    "sha256": hashlib.sha256(neutral_bytes).hexdigest(),
                },
                {
                    "input_id": "input.conversion.numbering-export-plan",
                    "locator": {"kind": "local_path", "path": str(numbering_plan)},
                    "kind": "resource",
                    "role": "numbering_export_plan",
                    "logical_path": "inputs/numbering-export-plan.json",
                    "media_type": MACHINE_NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
                    "size_bytes": len(numbering_bytes),
                    "sha256": hashlib.sha256(numbering_bytes).hexdigest(),
                },
            ],
            convert_staging,
            options={},
            label="Conversion",
        )

        validate_source = tmp_path / "private-validate.md"
        validate_source.write_text("# Private validation\n\nFullwidth digit: １\n", encoding="utf-8")
        validate_staging = tmp_path / "validate-staging"
        validate_staging.mkdir()
        validate_bytes = validate_source.read_bytes()
        validate_lifecycle = execute(
            "validate.markdown",
            [
                {
                    "input_id": "input.validation.source",
                    "locator": {"kind": "local_path", "path": str(validate_source)},
                    "kind": "document",
                    "role": "source",
                    "logical_path": "inputs/source.md",
                    "media_type": "text/markdown",
                    "size_bytes": len(validate_bytes),
                    "sha256": hashlib.sha256(validate_bytes).hexdigest(),
                }
            ],
            validate_staging,
            options={"enable_sensitive_word": False},
            label="Validation",
        )

        assert len([item for item in convert_lifecycle if item.get("method") == "task/progress"]) > 2
        assert len([item for item in validate_lifecycle if item.get("method") == "task/progress"]) > 2

        process_stdin.close()
        stdin_closed = True
        assert process.wait(timeout=30) == 0
        assert process.stderr.read() == b""
    finally:
        if not stdin_closed:
            process_stdin.close()
        if process.poll() is None:
            process.terminate()
            try:
                process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                process.kill()
                process.wait(timeout=5)

    validator = MachineContractValidator()
    for message in trace:
        validator.validate_message(message)
    validate_trace(trace, requires_terminal=True)
