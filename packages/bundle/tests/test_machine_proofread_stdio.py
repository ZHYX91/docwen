"""Proofreading through a real framed Machine session."""

from __future__ import annotations

import hashlib
import subprocess
import sys
import zipfile
from typing import BinaryIO, cast

import pytest
from docx import Document
from docx.oxml.ns import qn
from tools.validate_contracts import validate_trace

from docwen_application.conversion_contracts import DOCX_MEDIA_TYPE
from docwen_cli.machine.framing import FrameWriter, read_frame
from docwen_plugin_proofread.anchor_report import extract_occurrences_from_document_xml

from ._machine_stdio_e2e_support import _assert_bundle_files

pytestmark = [pytest.mark.e2e, pytest.mark.pr_gate]


@pytest.fixture
def machine(tmp_path, monkeypatch):
    monkeypatch.setenv("DOCWEN_RUNTIME_ROOT", str(tmp_path / "runtime"))
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

        def terminal():
            while True:
                message = read_frame(cast(BinaryIO, process.stdout))
                assert message is not None
                trace.append(message)
                if message.get("method") in {"task/completed", "task/failed", "task/cancelled"}:
                    return message

        yield exchange, terminal, trace
        validate_trace(trace, requires_terminal=True)
    finally:
        if process.stdin is not None:
            process.stdin.close()
        process.wait(timeout=30)
        assert process.stderr is not None
        assert process.stderr.read() == b""
        assert process.returncode == 0


def test_stdio_proofreads_generated_docx_without_exposing_private_intent(tmp_path, machine):
    exchange, terminal, trace = machine
    source = tmp_path / "source.md"
    source.write_text("# Review\n\n１ （\n", encoding="utf-8")
    data = source.read_bytes()
    output = tmp_path / "output"
    output.mkdir()
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
    message = terminal()
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


@pytest.mark.parametrize("text", ["A １ B", "A 1 B"], ids=["new-comment", "zero-findings"])
def test_stdio_proofreads_existing_docx(tmp_path, machine, text):
    exchange, terminal, _trace = machine
    source = tmp_path / "review.docx"
    document = Document()
    original_comment = document.add_comment(document.add_paragraph("Original").runs, "Keep this", author="Reviewer")
    document.add_paragraph(text)
    document.save(source)
    original = source.read_bytes()
    staging = tmp_path / "output"
    staging.mkdir()
    rules = {
        "enable_symbol_pairing": False,
        "enable_symbol_correction": True,
        "enable_typos_rule": False,
        "enable_sensitive_word": False,
    }
    discovery = exchange(2, "capability/list", {})
    capability = next(c for c in discovery["result"]["capabilities"] if c["capability_id"] == "validate.docx")
    assert capability["availability"] == "available"
    assert capability["options_schema"]["properties"] == {name: {"type": "boolean"} for name in rules}
    assert capability["options_schema"]["additionalProperties"] is False
    assert capability["options_schema"]["required"] == []
    params = {
        "capability_id": "validate.docx",
        "inputs": [
            {
                "input_id": "doc",
                "kind": "document",
                "role": "source",
                "logical_path": source.name,
                "locator": {"kind": "local_path", "path": str(source)},
                "media_type": DOCX_MEDIA_TYPE,
                "size_bytes": len(original),
                "sha256": hashlib.sha256(original).hexdigest(),
            }
        ],
        "output": {"staging_root": {"kind": "local_path", "path": str(staging)}, "staging_policy": "require_empty"},
        "options": rules,
    }
    for number, (options, code) in enumerate(
        [
            ({"proofread": {}}, "unsupported_options"),
            ({"enable_symbol_correction": 1}, "option_value_invalid"),
            (dict.fromkeys(rules, False), "proofread_no_rules_enabled"),
        ],
        start=3,
    ):
        rejected = exchange(number, "task/plan", {**params, "options": options})
        assert rejected["error"]["data"]["task_error"]["code"] == code
    assert list(staging.iterdir()) == []
    plan = exchange(6, "task/plan", params)["result"]
    assert plan["effective_options"] == rules
    exchange(7, "task/execute", {"plan_id": plan["plan_id"]})
    message = terminal()
    assert message["method"] == "task/completed", message
    bundle = message["params"]["bundle"]
    [artifact] = bundle["artifacts"]
    assert artifact["kind"] == "document" and artifact["media_type"] == DOCX_MEDIA_TYPE
    assert bundle["entries"] == [
        {"artifact_id": artifact["artifact_id"], "role": "primary", "ordinal": 0, "preferred": True}
    ]
    assert bundle["relations"] == []
    _assert_bundle_files(bundle, staging)
    final = staging / artifact["locator"]
    reopened = Document(final)
    assert [p.text for p in reopened.paragraphs] == ["Original", text]
    old, *added = reopened.comments
    assert (old.comment_id, old.author, old.text) == (original_comment.comment_id, "Reviewer", "Keep this")
    with zipfile.ZipFile(final) as package:
        anchors, _diagnostics = extract_occurrences_from_document_xml(package.read("word/document.xml"), 20, False)
    expected_anchors = [(str(old.comment_id), "Original")]
    if "１" in text:
        [comment] = added
        expected_anchors.append((str(comment.comment_id), "１"))
        assert "１" in comment.text
        [left] = reopened.paragraphs[1]._p.xpath("./w:r/w:t[text()='A ']")
        assert left.get(qn("xml:space")) == "preserve"
    else:
        assert added == []
    assert [(a.comment_id, a.covered_text) for a in anchors] == expected_anchors
    assert source.read_bytes() == original
