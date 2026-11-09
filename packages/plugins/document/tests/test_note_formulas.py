"""Native note formulas use the same parser and diagnostics as body formulas."""

from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

import pytest
from docx import Document
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from lxml import etree
from tests.support.config import FakeConfigView
from tests.support.execution import FakeExecutionContext
from tests.support.logging import FakePluginLogger
from tests.support.progress import FakeProgressSink
from tests.support.workspace import FakeWorkspaceHandle

from docwen_core.cancellation import CancellationToken
from docwen_core.formula.constants import OMML_NS
from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_plugin_document.shared.note_extraction import _extract_note_content
from docwen_plugin_document.to_markdown.converter import DocxToMarkdownConverter

pytestmark = pytest.mark.contract
_W = OMML_NS["w"]
_M = OMML_NS["m"]
_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_SQUARE = "<m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>"
_UNSUPPORTED = "<m:acc><m:e><m:r><m:t>x</m:t></m:r></m:e></m:acc>"


def _math(contents: str, *, display: bool = False) -> etree._Element:
    formula = f'<m:oMath xmlns:m="{_M}">{contents}</m:oMath>'
    if display:
        formula = f'<m:oMathPara xmlns:m="{_M}">{formula}</m:oMathPara>'
    return etree.fromstring(formula.encode())


def _text(parent: etree._Element, value: str) -> None:
    run = etree.SubElement(parent, f"{{{_W}}}r")
    text = etree.SubElement(run, f"{{{_W}}}t")
    text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text.text = value


def _document_with_note(
    path: Path,
    kind: str,
    *,
    layout: str = "inline",
    related_part: bool = True,
    filename: str | None = None,
    contents: str = _SQUARE,
) -> str:
    """Build independent Word XML; never create the source via our math writer."""
    document = Document()
    paragraph = document.add_paragraph("Body")
    marker = etree.SubElement(paragraph.add_run()._r, f"{{{_W}}}{kind}Reference")
    marker.set(f"{{{_W}}}id", "7")
    paragraph.add_run(".")

    root = etree.Element(f"{{{_W}}}{kind}s", nsmap={"w": _W, "m": _M})
    note = etree.SubElement(root, f"{{{_W}}}{kind}")
    note.set(f"{{{_W}}}id", "7")
    body = etree.SubElement(note, f"{{{_W}}}p")
    etree.SubElement(etree.SubElement(body, f"{{{_W}}}r"), f"{{{_W}}}{kind}Ref")
    _text(body, " ")
    if layout == "mixed":
        _text(body, "Before ")
    body.append(_math(contents, display=layout == "display"))
    if layout == "mixed":
        _text(body, " after ")
        body.append(_math("<m:r><m:t>y</m:t></m:r>"))
        _text(body, " done")

    source_part = "word/" + (filename or f"{kind}s.xml")
    blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    if related_part:
        part = Part(
            PackURI("/" + source_part),
            CT.WML_FOOTNOTES if kind == "footnote" else CT.WML_ENDNOTES,
            blob,
            document.part.package,
        )
        document.part.relate_to(part, RT.FOOTNOTES if kind == "footnote" else RT.ENDNOTES)
    document.save(str(path))
    if not related_part:
        with ZipFile(path, "a") as package:
            package.writestr(source_part, blob)
    return source_part


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
@pytest.mark.parametrize("related_part", [False, True], ids=["zip-fallback", "package-part"])
@pytest.mark.parametrize(
    ("layout", "expected"),
    [("inline", "$x^{2}$"), ("mixed", "Before $x^{2}$ after $y$ done"), ("display", "$$x^{2}$$")],
)
def test_native_note_formula_content_survives_the_production_converter(tmp_path, kind, related_part, layout, expected):
    source = tmp_path / "native-note.docx"
    _document_with_note(source, kind, layout=layout, related_part=related_part)
    result, markdown = _convert(source)

    key = "1" if kind == "footnote" else "endnote:1"
    assert f"Body[^{key}]." in markdown
    assert f"[^{key}]: {expected}" in markdown
    assert [item for item in result.diagnostics if "FORMULA" in item.code] == []


@pytest.mark.parametrize("kind", ["footnote", "endnote"])
@pytest.mark.parametrize("custom_part", [False, True])
def test_unsupported_note_formula_keeps_definition_and_reports_the_actual_part(tmp_path, kind, custom_part):
    source = tmp_path / "unsupported-note.docx"
    source_part = _document_with_note(
        source,
        kind,
        filename=f"authored-{kind}s.xml" if custom_part else None,
        contents=_UNSUPPORTED,
    )
    result, markdown = _convert(source)

    key = "1" if kind == "footnote" else "endnote:1"
    definition = next(line for line in markdown.splitlines() if line.startswith(f"[^{key}]: "))
    assert definition.removeprefix(f"[^{key}]: ").strip()
    [diagnostic] = [item for item in result.diagnostics if "FORMULA" in item.code]
    assert diagnostic.code == "DOCX2MD-FORMULA-SOURCE-FALLBACK"
    assert diagnostic.location == f"{source_part}:/w:{kind}s/w:{kind}/w:p/m:oMath"
    assert "unsupported" in diagnostic.message


@pytest.mark.parametrize("kind", ["footnote", "endnote"])
def test_note_walker_reads_only_effective_formulas_in_source_order(kind):
    note = etree.fromstring(
        f'<w:{kind} xmlns:w="{_W}" xmlns:m="{_M}" xmlns:mc="{_MC}"><w:p>'
        f'<w:r><w:{kind}Ref/></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>'
        '<w:r><w:t xml:space="preserve">Before </w:t></w:r>'
        f"<w:del><m:oMath>{_UNSUPPORTED}</m:oMath></w:del>"
        f"<w:ins><m:oMath>{_SQUARE}</m:oMath></w:ins>"
        '<w:r><w:t xml:space="preserve"> middle </w:t></w:r>'
        '<mc:AlternateContent><mc:Choice Requires="m"><m:oMath><m:r><m:t>y</m:t></m:r></m:oMath>'
        "</mc:Choice><mc:Fallback><m:oMath><m:r><m:t>duplicate</m:t></m:r></m:oMath>"
        "</mc:Fallback></mc:AlternateContent>"
        "<w:r><w:rPr><w:vanish/></w:rPr><w:t>hidden</w:t></w:r>"
        '<w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>'
        "<w:p><m:oMathPara><m:oMath><m:r><m:t>z</m:t></m:r></m:oMath></m:oMathPara></w:p>"
        f"</w:{kind}>"
    )
    original = etree.tostring(note)
    converter = DocxToMarkdownConverter()
    result = _extract_note_content(note, _W, f"{kind}Ref", math_renderer=converter._render_formula)

    assert result == "Before $x^{2}$ middle $y$ after\n\n$$z$$"
    assert converter._formula_diagnostics == []
    assert etree.tostring(note) == original


def test_math_between_note_marker_and_space_keeps_the_authored_space():
    note = etree.fromstring(
        f'<w:footnote xmlns:w="{_W}" xmlns:m="{_M}"><w:p><w:r><w:footnoteRef/></w:r>'
        f"<m:oMath>{_SQUARE}</m:oMath>"
        '<w:r><w:t xml:space="preserve"> </w:t></w:r></w:p></w:footnote>'
    )
    assert _extract_note_content(note, _W, "footnoteRef") == "$x^{2}$ "
