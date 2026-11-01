"""Plugin dispatch keeps admitted text format and source-role identity distinct."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pytest

from ._proofread_plugin_support import _build_fake_context

pytestmark = pytest.mark.contract


@pytest.mark.parametrize("source_index", [0, 1])
@pytest.mark.parametrize("source_format", ["txt", "markdown"])
def test_admitted_text_source_after_resource_keeps_report_identity(tmp_path, source_index, source_format):
    from docwen_core.models.file_ref import FileRef
    from docwen_plugin_proofread import ProofreadPlugin

    source = tmp_path / "authored.md"
    raw = b"Ordinary Markdown paragraph.\n"
    source.write_bytes(raw)
    resource = tmp_path / "not-the-document.bin"
    resource.write_bytes(b"resource bytes must not become source text")
    staging = tmp_path / "staging"
    staging.mkdir()
    # This is a narrow plugin-context test; real Application/Runtime/report
    # publication is covered by test_validate_execute_02.
    context = _build_fake_context(
        str(source), str(staging), source_format=source_format, target_format="markdown", action_name="validate"
    )
    document_ref = context.request.input_refs[0]
    document_ref.category = "markdown"
    refs = [FileRef(path=str(resource), format="binary", category="unknown", input_role="linked_resource")]
    refs.insert(source_index, document_ref)
    context.request.input_refs = refs
    result = ProofreadPlugin().convert(context)
    assert result.success
    [artifact] = result.artifacts
    report = json.loads(Path(artifact.staging_path).read_text(encoding="utf-8"))
    assert report["file"] == source.name
    assert report["source"]["content_sha256"] == hashlib.sha256(raw).hexdigest()
    assert artifact.metadata["source_format"] == source_format
    assert source.read_bytes() == raw


@pytest.mark.parametrize(("source_format", "category"), [("txt", "document"), ("html", "markdown")])
def test_category_fallback_does_not_broaden_unsupported_content(tmp_path, source_format, category):
    from docwen_plugin_proofread import ProofreadPlugin

    source = tmp_path / "source.md"
    source.write_text("text", encoding="utf-8")
    staging = tmp_path / "staging"
    staging.mkdir()
    context = _build_fake_context(
        str(source), str(staging), source_format=source_format, target_format="markdown", action_name="validate"
    )
    context.request.input_refs[0].category = category
    result = ProofreadPlugin().convert(context)
    assert not result.success and result.error is not None
    assert result.error.error_type == "unsupported_route"
    assert not result.artifacts and not list(staging.iterdir())
