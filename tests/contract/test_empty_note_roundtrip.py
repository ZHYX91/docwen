"""Cross-plugin acceptance: existing empty note definitions survive import."""

from __future__ import annotations

from pathlib import Path

import pytest
from docx import Document
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from lxml import etree

from docwen_core.cancellation import CancellationToken
from docwen_core.formula.constants import OMML_NS
from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_plugin_document.to_markdown.converter import DocxToMarkdownConverter
from tests.support.config import FakeConfigView
from tests.support.execution import FakeExecutionContext
from tests.support.logging import FakePluginLogger
from tests.support.progress import FakeProgressSink
from tests.support.workspace import FakeWorkspaceHandle

pytestmark = pytest.mark.contract
_W = OMML_NS["w"]


def _convert(source: Path):
    staging = source.parent / "staging"
    staging.mkdir()
    context = FakeExecutionContext(
        request=ConversionRequest(
            request_id="native-note-formulas",
            input_refs=[FileRef(path=str(source), format="docx", category="document")],
            target_format="md",
            options={"markdown_extensions": {"output": {"typed_endnotes": True}}},
            output_policy=OutputPolicy(),
        ),
        workspace=FakeWorkspaceHandle(str(source), str(staging)),
        config=FakeConfigView(),
        progress=FakeProgressSink(),
        cancellation=CancellationToken().view(),
        logger=FakePluginLogger(),
        numbering_registry=None,
    )
    result = DocxToMarkdownConverter().convert(context)
    assert result.success, result.error
    return result, Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")


@pytest.mark.parametrize("kind", ["footnote", "endnote"])
def test_present_empty_note_definition_is_not_reported_as_unreadable(tmp_path, kind):
    from docwen_plugin_markdown.mistune_extensions import parse_markdown_text
    from docwen_plugin_markdown.to_docx.notes import extract_notes_from_ast, normalize_note_syntax

    source = tmp_path / "empty-note.docx"
    document = Document()
    marker = etree.SubElement(document.add_paragraph("Body").add_run()._r, f"{{{_W}}}{kind}Reference")
    marker.set(f"{{{_W}}}id", "7")
    root = etree.Element(f"{{{_W}}}{kind}s", nsmap={"w": _W})
    note = etree.SubElement(root, f"{{{_W}}}{kind}")
    note.set(f"{{{_W}}}id", "7")
    etree.SubElement(etree.SubElement(etree.SubElement(note, f"{{{_W}}}p"), f"{{{_W}}}r"), f"{{{_W}}}{kind}Ref")
    part = Part(
        PackURI(f"/word/{kind}s.xml"),
        CT.WML_FOOTNOTES if kind == "footnote" else CT.WML_ENDNOTES,
        etree.tostring(root),
        document.part.package,
    )
    document.part.relate_to(part, RT.FOOTNOTES if kind == "footnote" else RT.ENDNOTES)
    document.save(source)
    result, markdown = _convert(source)
    key = "1" if kind == "footnote" else "endnote:1"
    assert f"[^{key}]:" in markdown
    assert not [item for item in result.diagnostics if "LOSS" in item.code]
    # The emitted blank definition must remain valid for the production importer.
    ast, notes = extract_notes_from_ast(parse_markdown_text(normalize_note_syntax(markdown)))
    assert ast
    reference = notes.create_footnote_ref_run("1") if kind == "footnote" else notes.create_endnote_ref_run("1")
    assert reference is not None
