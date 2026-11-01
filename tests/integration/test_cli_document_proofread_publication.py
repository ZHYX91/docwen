"""Real private DOCX proofreading and final publication; Office is an explicit double."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET
from zipfile import ZIP_STORED, ZipFile

import pytest

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_OFFICE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
_TEXT = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
_BODY = "Ordinary Markdown paragraph."


def _docx_facts(path: Path) -> tuple[list[str], list[str]]:
    with ZipFile(path) as package:
        document = ET.fromstring(package.read("word/document.xml"))
        body = document.find(f"{_W}body")
        assert body is not None
        paragraphs = [
            "".join(node.text or "" for node in paragraph.iter(f"{_W}t")) for paragraph in body.iter(f"{_W}p")
        ]
        comments = []
        if "word/comments.xml" in package.namelist():
            comments_root = ET.fromstring(package.read("word/comments.xml"))
            comments = ["".join(node.text or "" for node in comment.iter(f"{_W}t")) for comment in comments_root]
        return paragraphs, comments


def _controller(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    from docwen_application.controller import ApplicationController
    from docwen_bundle.config_port import ConfigPortAdapter
    from docwen_bundle.runtime_factory import create_runtime_port
    from docwen_runtime.config import ConfigLoader

    monkeypatch.setenv("DOCWEN_RUNTIME_ROOT", str(tmp_path / "runtime"))
    loader = ConfigLoader(
        user_dir=tmp_path / "config",
        runtime_overrides={
            "logger": {"console_enable": False},
            "proofread": {
                "engine": {
                    "enable_symbol_pairing": False,
                    "enable_symbol_correction": False,
                    "enable_typos_rule": True,
                    "enable_sensitive_word": False,
                },
                "typos": {"entries": {"Expected": ["Ordinary"]}},
            },
        },
    )
    controller = ApplicationController(
        runtime_port=create_runtime_port(config_loader=loader), config_port=ConfigPortAdapter(loader)
    )
    controller.start()
    return controller


def _allow_office_double(monkeypatch: pytest.MonkeyPatch) -> None:
    from docwen_runtime import capabilities

    original = capabilities._office_gate_status

    def status(gate_id: str) -> dict[str, Any]:
        if gate_id != "external_office.word":
            return original(gate_id)
        return {
            "id": gate_id,
            "kind": "external_office",
            "label": "Explicit test bridge",
            "available": True,
            "reason": None,
            "providers": ["test-bridge"],
        }

    monkeypatch.setattr(capabilities, "_office_gate_status", status)


@pytest.mark.parametrize(("target", "checks"), [("docx", []), ("docx", ["--check", "none"]), ("odt", [])])
def test_cli_proofread_publishes_checked_document_once(tmp_path, monkeypatch, capsys, target, checks):
    from docwen_cli.main import main
    from docwen_core.office_bridge import BridgeResult
    from docwen_plugin_document.to_document.converter import SmartDocConverter
    from docwen_runtime.output.finalizer import OutputFinalizer

    source = tmp_path / "source.md"
    source.write_bytes(f"# Heading\n\n{_BODY}\n".encode())
    original_source = source.read_bytes()
    public = tmp_path / "published"
    public.mkdir()
    controller = _controller(tmp_path, monkeypatch)
    _allow_office_double(monkeypatch)
    public_results = []
    private_roots = []
    bridge_facts = []
    original_finalize = OutputFinalizer.finalize

    def finalize(self, *args, **kwargs):
        policy = kwargs["policy"]
        result = original_finalize(self, *args, **kwargs)
        parent = Path(policy.output_dir or "").resolve()
        if parent == public.resolve():
            public_results.append(result)
        else:
            private_roots.append(parent)
        return result

    def bridge(self, *, context, input_path, source, target, output_path, cancel=None):
        assert source == "docx" and target == "odt"
        if cancel is not None:
            cancel.check()
        facts = _docx_facts(Path(input_path))
        bridge_facts.append(facts)
        assert facts[0].count(_BODY) == 1
        assert len(facts[1]) == 1 and "Ordinary" in facts[1][0]
        # This is an ODF-shaped bridge double, not an installed Office host.
        content = ET.Element(f"{{{_OFFICE}}}document-content", {f"{{{_OFFICE}}}version": "1.2"})
        body = ET.SubElement(content, f"{{{_OFFICE}}}body")
        text = ET.SubElement(body, f"{{{_OFFICE}}}text")
        for value in facts[0]:
            ET.SubElement(text, f"{{{_TEXT}}}p").text = value
        with ZipFile(output_path, "w") as package:
            package.writestr("mimetype", "application/vnd.oasis.opendocument.text", compress_type=ZIP_STORED)
            package.writestr("content.xml", ET.tostring(content, encoding="utf-8", xml_declaration=True))
            package.writestr(
                "META-INF/manifest.xml",
                '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" '
                'manifest:version="1.2"><manifest:file-entry manifest:full-path="/" '
                'manifest:media-type="application/vnd.oasis.opendocument.text"/>'
                '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>'
                "</manifest:manifest>",
            )
        return BridgeResult(success=True, output_path=output_path, backend="explicit-test-bridge")

    monkeypatch.setattr(OutputFinalizer, "finalize", finalize)
    monkeypatch.setattr(SmartDocConverter, "_convert_one_hop", bridge)
    try:
        code = main(
            ["convert", str(source), "--to", target, "--output-dir", str(public), "--proofread", *checks, "--json"],
            controller=controller,
        )
        payload = json.loads(capsys.readouterr().out)
    finally:
        controller.stop()
    assert code == 0 and payload["success"], payload
    [committed] = public_results
    assert committed.success
    output = Path(payload["data"]["output"])
    assert output == Path(committed.artifacts[0].staging_path)
    if target == "docx":
        paragraphs, comments = _docx_facts(output)
        assert paragraphs.count(_BODY) == 1
        assert len(comments) == (0 if checks else 1)
        assert bridge_facts == []
    else:
        with ZipFile(output) as package:
            assert package.read("mimetype") == b"application/vnd.oasis.opendocument.text"
            content = ET.fromstring(package.read("content.xml"))
        assert ["".join(node.itertext()) for node in content.iter(f"{{{_TEXT}}}p")] == bridge_facts[0][0]
    assert all(not path.exists() for path in private_roots)
    assert source.read_bytes() == original_source


@pytest.mark.parametrize("outcome", ["failure", "cancel"])
def test_real_checked_docx_failure_or_cancellation_never_bridges_or_publishes(tmp_path, monkeypatch, capsys, outcome):
    from docwen_cli.main import main
    from docwen_core.models.result import ConversionErrorInfo, ConversionResult
    from docwen_plugin_document.to_document.converter import SmartDocConverter
    from docwen_plugin_proofread.docx_validator import DocxValidator

    source = tmp_path / "source.md"
    source.write_bytes(f"# Heading\n\n{_BODY}\n".encode())
    original_source = source.read_bytes()
    public = tmp_path / "published"
    public.mkdir()
    controller = _controller(tmp_path, monkeypatch)
    _allow_office_double(monkeypatch)
    root_task_ids = []
    checked_facts = []
    original_prepare = controller.prepare_execution_cancellation
    original_convert = DocxValidator.convert

    def prepare(request, *, batch=False, on_reserved=None):
        root_task_ids.append(request.request_id)
        return original_prepare(request, batch=batch, on_reserved=on_reserved)

    def convert(self, context):
        result = original_convert(self, context)
        assert result.success
        checked_facts.append(_docx_facts(Path(result.artifacts[0].staging_path)))
        if outcome == "cancel":
            controller.cancel(root_task_ids[0])
            return result
        return ConversionResult(
            task_id=context.request.request_id,
            success=False,
            error=ConversionErrorInfo(error_type="conversion_failed", message="Injected after real checking"),
        )

    def forbidden_bridge(*_args, **_kwargs):
        pytest.fail("Failed or cancelled proofreading must not invoke the external bridge")

    monkeypatch.setattr(controller, "prepare_execution_cancellation", prepare)
    monkeypatch.setattr(DocxValidator, "convert", convert)
    monkeypatch.setattr(SmartDocConverter, "_convert_one_hop", forbidden_bridge)
    try:
        code = main(
            ["convert", str(source), "--to", "odt", "--output-dir", str(public), "--proofread", "--json"],
            controller=controller,
        )
        payload = json.loads(capsys.readouterr().out)
    finally:
        controller.stop()
    assert code == (130 if outcome == "cancel" else 1), payload
    assert payload["success"] is False
    assert checked_facts[0][0].count(_BODY) == 1
    assert len(checked_facts[0][1]) == 1 and "Ordinary" in checked_facts[0][1][0]
    assert not list(public.iterdir())
    assert source.read_bytes() == original_source


@pytest.mark.parametrize("target", ["doc", "odt", "rtf", "wps"])
def test_missing_legacy_bridge_is_a_capability_failure_after_cli_admission(tmp_path, monkeypatch, capsys, target):
    from docwen_cli.main import main
    from docwen_runtime import capabilities

    source = tmp_path / "source.md"
    source.write_bytes(f"# Heading\n\n{_BODY}\n".encode())
    public = tmp_path / "published"
    public.mkdir()
    controller = _controller(tmp_path, monkeypatch)
    original = capabilities._office_gate_status

    def unavailable(gate_id):
        status = original(gate_id)
        if gate_id == "external_office.word":
            return {**status, "available": False, "providers": [], "reason": "no_compatible_backend"}
        return status

    monkeypatch.setattr(capabilities, "_office_gate_status", unavailable)
    try:
        code = main(
            ["convert", str(source), "--to", target, "--output-dir", str(public), "--proofread", "--json"],
            controller=controller,
        )
        payload = json.loads(capsys.readouterr().out)
    finally:
        controller.stop()
    assert code == 6 and payload["error"]["code"] == "capability_unavailable", payload
    assert not list(public.iterdir())
    assert source.read_bytes() == f"# Heading\n\n{_BODY}\n".encode()
