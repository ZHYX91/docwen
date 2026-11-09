"""Real framed Machine wire exposes source-only post-render proofreading."""

from __future__ import annotations

import hashlib
import subprocess
import sys
import zipfile
from typing import BinaryIO, cast

import pytest
from tools.validate_contracts import validate_trace

from docwen_cli.machine.framing import FrameWriter, read_frame

pytestmark = [pytest.mark.e2e, pytest.mark.pr_gate]


def test_stdio_proofreads_generated_docx_without_exposing_private_intent(tmp_path, monkeypatch):
    monkeypatch.setenv("DOCWEN_RUNTIME_ROOT", str(tmp_path / "runtime"))
    source = tmp_path / "source.md"
    source.write_text("# Review\n\n１ （\n", encoding="utf-8")
    data = source.read_bytes()
    output = tmp_path / "output"
    output.mkdir()
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
    trace = []
    try:
        assert process.stdin is not None and process.stdout is not None
        writer = FrameWriter(cast(BinaryIO, process.stdin))

        def exchange(number, method, params):
            message = {"jsonrpc": "2.0", "id": number, "method": method, "params": params}
            trace.append(message)
            writer.write(message)
            response = read_frame(cast(BinaryIO, process.stdout))
            assert response is not None
            trace.append(response)
            assert response.get("id") == number
            return response

        exchange(
            1,
            "initialize",
            {
                "protocol": {"name": "docwen.machine", "major": 2, "minor": 0},
                "client": {"name": "proofread-test", "version": "1"},
                "features": {"progress": True, "cancellation": True},
            },
        )
        discovery = exchange(2, "capability/list", {})
        capabilities = {item["capability_id"]: item for item in discovery["result"]["capabilities"]}
        properties = capabilities["convert.markdown_source.to_docx"]["options_schema"]["properties"]
        assert set(properties["proofread"]["properties"]) == {
            "enable_symbol_pairing",
            "enable_symbol_correction",
            "enable_typos_rule",
            "enable_sensitive_word",
        }
        assert "proofread" not in capabilities["convert.markdown.to_docx"]["options_schema"]["properties"]
        params = {
            "capability_id": "convert.markdown_source.to_docx",
            "inputs": [
                {
                    "input_id": "source",
                    "kind": "document",
                    "role": "source",
                    "logical_path": "source.md",
                    "locator": {"kind": "local_path", "path": str(source)},
                    "media_type": "text/markdown",
                    "size_bytes": len(data),
                    "sha256": hashlib.sha256(data).hexdigest(),
                }
            ],
            "output": {"staging_root": {"kind": "local_path", "path": str(output)}, "staging_policy": "require_empty"},
            "options": {
                "proofread": {
                    "enable_symbol_pairing": True,
                    "enable_symbol_correction": True,
                    "enable_typos_rule": False,
                    "enable_sensitive_word": False,
                }
            },
        }
        bad = exchange(3, "task/plan", {**params, "options": {"proofread": {"enable_symbol_pairing": 1}}})
        assert bad["error"]["data"]["task_error"]["code"] == "option_value_invalid"
        plan = exchange(4, "task/plan", params)
        assert plan["result"]["effective_options"]["proofread"] == params["options"]["proofread"]
        exchange(5, "task/execute", {"plan_id": plan["result"]["plan_id"]})
        while True:
            message = read_frame(cast(BinaryIO, process.stdout))
            assert message is not None
            trace.append(message)
            if message.get("method") in {"task/completed", "task/failed", "task/cancelled"}:
                break
        assert message["method"] == "task/completed", message
        [artifact] = message["params"]["bundle"]["artifacts"]
        final = output / artifact["locator"]
        assert hashlib.sha256(final.read_bytes()).hexdigest() == artifact["sha256"]
        with zipfile.ZipFile(final) as package:
            comments = package.read("word/comments.xml").decode()
            assert "１" in comments and "（" in comments
        assert source.read_bytes() == data
        assert [p for p in output.rglob("*") if p.is_file()] == [final]
        assert "_postprocess_proofread" not in str(trace)
        validate_trace(trace, requires_terminal=True)
    finally:
        if process.stdin is not None:
            process.stdin.close()
        process.wait(timeout=30)
        assert process.stderr is not None
        assert process.stderr.read() == b""
        assert process.returncode == 0
