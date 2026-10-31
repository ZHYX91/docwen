"""Admitted Markdown workflows through inspection and real report publication."""

from __future__ import annotations

import pytest

pytestmark = [pytest.mark.integration, pytest.mark.release_gate]


@pytest.mark.integration
@pytest.mark.parametrize(
    ("filename", "text", "detected", "expected_offsets"),
    [
        ("plain.md", "Ordinary Markdown paragraph.\n", "txt", [0]),
        ("inline.md", "A paragraph with `inline code`.\n", "txt", []),
        ("protected.md", "Ordinary `Ordinary`.\n", "txt", [0]),
        ("list.md", "- Ordinary Markdown paragraph.\n", "txt", [2]),
        ("heading.md", "# Ordinary\n", "markdown", [2]),
        ("plain.txt", "Ordinary Markdown paragraph.\n", "txt", [0]),
    ],
)
def test_admitted_markdown_workflow_reaches_real_report(
    tmp_path, monkeypatch, capsys, filename, text, detected, expected_offsets
):
    import hashlib
    import json

    from docwen_bundle.config_port import ConfigPortAdapter
    from docwen_bundle.runtime_factory import create_runtime_port
    from docwen_cli.main import main
    from docwen_core.detection import inspect_file
    from docwen_runtime.config import ConfigLoader

    source = tmp_path / filename
    source.write_bytes(text.encode("utf-8"))
    original = source.read_bytes()
    report_path = tmp_path / "report.json"
    monkeypatch.setenv("DOCWEN_RUNTIME_ROOT", str(tmp_path / "runtime"))
    loader = ConfigLoader(
        user_dir=tmp_path / "config",
        runtime_overrides={
            "logger": {"console_enable": False},
            "proofread": {"typos": {"entries": {"Expected": ["Ordinary"]}}},
        },
    )
    inspection = inspect_file(str(source))
    assert inspection.detected_format == detected
    assert inspection.workflow_category == "markdown"
    assert inspection.may_execute and not inspection.requires_explicit_acceptance
    code = main(
        ["validate", str(source), "--report", str(report_path), "--check", "typo", "--json"],
        runtime_port_factory=lambda: create_runtime_port(config_loader=loader),
        config_port_factory=lambda: ConfigPortAdapter(loader),
    )
    payload = json.loads(capsys.readouterr().out)
    assert code == 0, payload
    assert payload["success"]
    # Inspect final published bytes independently of presenter embedding.
    report = json.loads(report_path.read_text(encoding="utf-8"))
    assert report == payload["data"]["details"]["proofread"]
    assert report["schema"] == "docwen.proofread_report.v2"
    assert report["file"] == filename
    assert report["source"]["content_sha256"] == hashlib.sha256(original).hexdigest()
    assert report["location_contract"]["coordinate_system"] == "unicode_code_point"
    assert [issue["range"]["start"]["offset"] for issue in report["issues"]] == expected_offsets
    for issue, offset in zip(report["issues"], expected_offsets, strict=True):
        assert issue["matched_text"] == "Ordinary"
        assert issue["range"]["end"]["offset"] == offset + len("Ordinary")
        assert text[offset : offset + len("Ordinary")] == issue["matched_text"]
    assert source.read_bytes() == original


@pytest.mark.integration
@pytest.mark.parametrize(
    "body",
    ["<!doctype html><html><body><h1>Ordinary</h1></body></html>\n", "Name,Amount\nAlice,7\nBob,9\n"],
)
def test_wrong_family_markdown_suffix_still_requires_explicit_acceptance(tmp_path, capsys, body):
    import json

    from docwen_cli.main import main
    from docwen_core.detection import inspect_file

    source = tmp_path / "wrong.md"
    source.write_bytes(body.encode("utf-8"))
    original = source.read_bytes()
    report = tmp_path / "must-not-publish.json"
    inspection = inspect_file(str(source))
    assert inspection.workflow_category != "markdown"
    assert inspection.requires_explicit_acceptance
    code = main(["validate", str(source), "--report", str(report), "--json"])
    payload = json.loads(capsys.readouterr().out)
    assert code == 2 and payload["error"]["code"] == "file_format_confirmation_required"
    assert not report.exists()
    assert source.read_bytes() == original
