"""Real Machine conversion renders first and comments only the generated DOCX."""

from __future__ import annotations

import hashlib
import zipfile
from pathlib import Path

import pytest
from docx import Document
from lxml import etree
from PIL import Image

from docwen_application.controller import ApplicationController
from docwen_application.conversion_contracts import ConversionPlanRequest, LocalInputHandle, StagingOutputTarget
from docwen_application.conversion_service import ConversionService
from docwen_bundle.config_port import ConfigPortAdapter
from docwen_bundle.runtime_factory import create_runtime_port
from docwen_runtime.config import ConfigLoader
from docwen_runtime.output.artifact_bundle import ArtifactBundleCommitter

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]
RULES = ("enable_symbol_pairing", "enable_symbol_correction", "enable_typos_rule", "enable_sensitive_word")


def _handle(path, role, media, logical):
    data = path.read_bytes()
    return LocalInputHandle(
        role,
        str(path),
        media,
        len(data),
        hashlib.sha256(data).hexdigest(),
        "document" if role == "source" else "resource",
        role,
        logical,
    )


@pytest.mark.parametrize("selection", [None, "none", "all", *RULES, "config", "override", "partial", "config_off"])
def test_generated_docx_comments_and_atomic_bundle(tmp_path: Path, monkeypatch, selection):
    monkeypatch.setenv("DOCWEN_RUNTIME_ROOT", str(tmp_path / "runtime"))
    source = tmp_path / "authored.md"
    source.write_text("# Review\n\n才料 秘密 １ （\n\n![figure](figure.png)\n", encoding="utf-8")
    original = source.read_bytes()
    image = tmp_path / "figure.png"
    Image.new("RGB", (8, 8), "red").save(image)
    output = tmp_path / "output"
    output.mkdir()
    loader = ConfigLoader(
        user_dir=tmp_path / "config",
        runtime_overrides={
            "logger": {"console_enable": False},
            "proofread": {
                "engine": dict.fromkeys(RULES, selection not in ("override", "config_off")),
                "typos": {"entries": {"材料": ["才料"]}},
                "sensitive_words": {"entries": {"秘密": []}},
            },
        },
    )
    controller = ApplicationController(
        runtime_port=create_runtime_port(config_loader=loader), config_port=ConfigPortAdapter(loader)
    )
    controller.start()
    service = ConversionService(controller, ArtifactBundleCommitter())
    options = {"locale": "zh_CN", "add_numbering": False}
    if selection is not None:
        options["proofread"] = (
            {}
            if selection in ("config", "config_off")
            else {key: selection in ("all", "override") or selection == key for key in RULES}
        )
    if selection == "partial":
        options["proofread"] = {"enable_typos_rule": False}
    request = ConversionPlanRequest(
        "convert.markdown_source.to_docx",
        (
            _handle(image, "linked_resource", "image/png", "notes/figure.png"),
            _handle(source, "source", "text/markdown", "notes/authored.md"),
        ),
        StagingOutputTarget(str(output)),
        options,
    )
    plan = service.plan(request)
    outcome = service.execute_accepted(service.accept(plan.plan_id))
    assert outcome.state == "completed", (outcome.error, outcome.diagnostics)
    assert outcome.bundle is not None
    [artifact] = outcome.bundle.artifacts
    assert artifact.kind == "document"
    assert len(outcome.bundle.entries) == 1 and not outcome.bundle.relations
    final = output / artifact.locator
    assert final.stat().st_size == artifact.size_bytes
    assert hashlib.sha256(final.read_bytes()).hexdigest() == artifact.sha256
    assert [p for p in output.rglob("*") if p.is_file()] == [final]
    assert any("才料 秘密 １ （" in p.text for p in Document(str(final)).paragraphs)
    with zipfile.ZipFile(final) as package:
        assert any(name.startswith("word/media/") for name in package.namelist())
        comments = (
            etree.fromstring(package.read("word/comments.xml")) if "word/comments.xml" in package.namelist() else []
        )
        texts = ["".join(str(text) for text in comment.itertext()) for comment in comments]
    expected = (
        0 if selection in (None, "none", "config_off") else 4 if selection in ("all", "config", "override") else 1
    )
    if selection == "partial":
        expected = 3
        assert not any("才料" in text for text in texts)
    assert len(texts) == expected, texts
    if selection in ("all", "config", "override"):
        assert any("才料 → 材料" in text for text in texts)
        assert any("１ → 1" in text for text in texts)
        assert any("秘密" in text for text in texts)
        assert any("（" in text for text in texts)
    codes = {d.code for d in outcome.diagnostics}
    assert ("POSTPROCESS_PROOFREAD_APPLIED" in codes) == (expected > 0)
    assert ("POSTPROCESS_PROOFREAD_SKIPPED" in codes) == (selection in ("none", "config_off"))
    assert source.read_bytes() == original
