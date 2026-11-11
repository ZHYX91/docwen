"""Saved body output retains authored XML formatting and semantic overrides."""

from __future__ import annotations

from hashlib import sha256
from pathlib import Path

import pytest
from docx import Document
from docx.document import Document as DocxDocument
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.oxml.text.run import CT_R
from docx.shared import Pt, RGBColor
from docx.styles.style import CharacterStyle, ParagraphStyle
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from lxml import etree

from docwen_plugin_markdown.template_utils import extract_body_paragraph_format, extract_body_run_format

from ._md_to_docx_support import MdToDocxConverter, make_context

pytestmark = pytest.mark.contract

_EXTENSION = "{urn:docwen:test:format-extension}format"


def _add(parent: etree._Element, name: str, **attributes: str) -> etree._Element:
    child = OxmlElement(f"w:{name}")
    for key, value in attributes.items():
        child.set(qn(f"w:{key}"), value)
    parent.append(child)
    return child


def _convert(tmp_path: Path, template: DocxDocument, markdown: str) -> DocxDocument:
    template_path = tmp_path / "template.docx"
    template.save(str(template_path))
    before = sha256(template_path.read_bytes()).hexdigest()
    source = tmp_path / "source.md"
    source.write_text(markdown, encoding="utf-8")
    context, _ = make_context(str(source), options={"template_name": str(template_path)})
    result = MdToDocxConverter().convert(context)
    assert result.success, result.error
    assert sha256(template_path.read_bytes()).hexdigest() == before
    return Document(result.artifacts[0].staging_path)


def _runs(paragraph: Paragraph) -> list[Run]:
    runs = []
    for element in paragraph._p.iter(qn("w:r")):
        assert isinstance(element, CT_R)
        run = Run(element, paragraph)
        if run.text:
            runs.append(run)
    return runs


def _required(parent: etree._Element | None, path: str) -> etree._Element:
    assert parent is not None
    child = parent.find(path)
    assert child is not None
    return child


def _run_property(run: Run, name: str) -> etree._Element:
    return _required(run._r.rPr, qn(f"w:{name}"))


def test_full_direct_anchor_format_reaches_body_without_changing_semantic_blocks(tmp_path: Path) -> None:
    template = Document()
    style = template.styles.add_style("Authored Body", WD_STYLE_TYPE.PARAGRAPH)
    assert isinstance(style, ParagraphStyle)
    style.base_style = template.styles["Normal"]
    anchor = template.add_paragraph("{{body}}", style=style)
    p_pr = anchor._p.get_or_add_pPr()
    borders = _add(p_pr, "pBdr")
    _add(borders, "bottom", val="single", sz="6", color="123456")
    _add(p_pr, "shd", val="clear", fill="EEEEDD")
    tabs = _add(p_pr, "tabs")
    _add(tabs, "tab", val="right", leader="dot", pos="4320")
    _add(p_pr, "spacing", beforeLines="120", afterAutospacing="1", line="360", lineRule="exact")
    _add(p_pr, "contextualSpacing")
    _add(p_pr, "textDirection", val="lrTb")
    mark = _add(p_pr, "rPr")
    _add(mark, "rFonts", ascii="ParagraphMarkFont", eastAsia="段落标记字体")
    _add(mark, "sz", val="33")
    p_pr.append(etree.Element(_EXTENSION, feature="paragraph"))
    p_pr.append(etree.Element(_EXTENSION, feature="second paragraph feature"))
    p_pr.append(etree.Comment("Paragraph format annotation"))
    p_pr.set(_EXTENSION, "paragraph attribute")
    r_pr = anchor.runs[0]._r.get_or_add_rPr()
    _add(r_pr, "smallCaps")
    _add(r_pr, "spacing", val="18")
    _add(r_pr, "highlight", val="yellow")
    _add(r_pr, "shd", val="clear", fill="ABCDEF")
    _add(r_pr, "lang", val="en-GB", eastAsia="zh-TW", bidi="ar-SA")
    _add(r_pr, "vertAlign", val="baseline")
    r_pr.append(etree.Element(_EXTENSION, feature="character"))
    r_pr.append(etree.Element(_EXTENSION, feature="second character feature"))
    r_pr.append(etree.Comment("Character format annotation"))
    r_pr.set(_EXTENSION, "character attribute")

    output = _convert(
        tmp_path,
        template,
        "Plain **strong** [link](https://example.com) `code` H~2~O.\n\n"
        "Second body.\n\n> Quote.\n\n> <p>HTML quote.</p>\n\n- Item.\n\n```\nCode block.\n```\n\n"
        "| Header |\n| --- |\n| Cell |\n",
    )
    ordinary = [p for p in output.paragraphs if p.text.startswith(("Plain ", "Second body."))]
    assert len(ordinary) == 2
    for paragraph in ordinary:
        properties = paragraph._p.pPr
        assert properties is not None
        for name in ("pBdr", "shd", "tabs", "spacing", "contextualSpacing", "textDirection", "rPr"):
            source = p_pr.find(qn(f"w:{name}"))
            actual = properties.find(qn(f"w:{name}"))
            assert source is not None and actual is not None
            assert etree.tostring(actual, method="c14n", exclusive=True) == etree.tostring(
                source, method="c14n", exclusive=True
            )
        assert _required(properties, _EXTENSION).get("feature") == "paragraph"
        assert len(properties.findall(_EXTENSION)) == 2
        assert properties.get(_EXTENSION) == "paragraph attribute"
        assert [node.text for node in properties if isinstance(node, etree._Comment)] == ["Paragraph format annotation"]
    body_runs = _runs(ordinary[0])
    for text in ("Plain ", "strong", "link"):
        run = next(run for run in body_runs if run.text == text)
        assert _run_property(run, "smallCaps") is not None
        assert _run_property(run, "spacing").get(qn("w:val")) == "18"
        assert _run_property(run, "highlight").get(qn("w:val")) == "yellow"
        assert _run_property(run, "lang").get(qn("w:eastAsia")) == "zh-TW"
        assert _required(run._r.rPr, _EXTENSION).get("feature") == "character"
        assert len(_required(run._r, qn("w:rPr")).findall(_EXTENSION)) == 2
        assert run._r.rPr is not None and run._r.rPr.get(_EXTENSION) == "character attribute"
        assert [node.text for node in run._r.rPr if isinstance(node, etree._Comment)] == ["Character format annotation"]
    assert next(run for run in body_runs if run.text == "strong").bold is True
    assert next(run for run in body_runs if run.text == "link").style.style_id == "Hyperlink"
    code = next(run for run in body_runs if run.text == "code")
    assert code.style.style_id == "DocWenInlineCode"
    assert code._r.rPr is not None
    assert code._r.rPr.find(qn("w:shd")) is None  # The code style owns its shading.
    assert _run_property(next(run for run in body_runs if run.text == "2"), "vertAlign").get(qn("w:val")) == (
        "subscript"
    )
    for paragraph in [*output.paragraphs, *output.tables[0].cell(1, 0).paragraphs]:
        if any(paragraph._p is body._p for body in ordinary):
            continue
        properties = paragraph._p.pPr
        assert properties is None or properties.find(_EXTENSION) is None
        assert all(run._r.rPr is None or run._r.rPr.find(_EXTENSION) is None for run in _runs(paragraph))


def test_projection_preserves_paragraph_marks_but_does_not_clone_source_identity(tmp_path: Path) -> None:
    document = Document()
    anchor = document.add_paragraph("{{body}}")
    anchor._p.set(qn("w14:paraId"), "12345678")
    p_pr = anchor._p.get_or_add_pPr()
    _add(_add(p_pr, "numPr"), "numId", val="1")
    _add(p_pr, "sectPr")
    _add(p_pr, "divId", val="12")
    _add(_add(p_pr, "pPrChange", id="5", author="Template"), "pPr")
    mark = _add(p_pr, "rPr")
    _add(mark, "sz", val="35")
    _add(mark, "lang", val="ja-JP")
    _add(mark, "ins", id="6", author="Template")
    _add(_add(mark, "rPrChange", id="7", author="Template"), "rPr")
    r_pr = anchor.runs[0]._r.get_or_add_rPr()
    _add(r_pr, "highlight", val="yellow")
    _add(_add(r_pr, "rPrChange", id="8", author="Template"), "rPr")
    r_pr.set(qn("w:rsidRPr"), "12345678")
    paragraph_format = extract_body_paragraph_format(document)
    run_format = extract_body_run_format(document)
    assert paragraph_format is not None and run_format is not None
    empty = document.add_paragraph()
    paragraph_format.apply_to(empty)
    target = document.add_paragraph("Generated", style="Heading 2")
    target_p_pr = target._p.get_or_add_pPr()
    _add(_add(target_p_pr, "numPr"), "numId", val="2")
    _add(_add(target_p_pr, "rPr"), "sz", val="44")
    target.runs[0].font.highlight_color = WD_COLOR_INDEX.BRIGHT_GREEN
    paragraph_format.apply_to(target)
    run_format.apply_to(target)
    path = tmp_path / "projected.docx"
    document.save(str(path))
    output = Document(str(path))
    empty, target = output.paragraphs[1:]
    assert _required(empty._p.pPr, f"{qn('w:rPr')}/{qn('w:sz')}").get(qn("w:val")) == "35"
    assert _required(target._p.pPr, f"{qn('w:rPr')}/{qn('w:sz')}").get(qn("w:val")) == "44"
    assert _required(target._p.pPr, f"{qn('w:rPr')}/{qn('w:lang')}").get(qn("w:val")) == "ja-JP"
    assert target.style is not None and target.style.style_id == "Heading2"
    assert _required(target._p.pPr, f"{qn('w:numPr')}/{qn('w:numId')}").get(qn("w:val")) == "2"
    for paragraph in (empty, target):
        assert paragraph._p.get(qn("w14:paraId")) is None
        for name in ("sectPr", "pPrChange", "rPrChange", "ins", "divId"):
            assert not list(paragraph._p.iter(qn(f"w:{name}")))
    assert _run_property(target.runs[0], "highlight").get(qn("w:val")) == "green"
    assert target.runs[0]._r.rPr is not None and target.runs[0]._r.rPr.get(qn("w:rsidRPr")) is None


def test_authored_character_style_survives_plain_and_merged_body_but_semantic_styles_win(tmp_path: Path) -> None:
    template = Document()
    body_style = template.styles.add_style("Body Base", WD_STYLE_TYPE.PARAGRAPH)
    assert isinstance(body_style, ParagraphStyle)
    body_style.font.name = "Body Font"
    body_style.font.size = Pt(12)
    heading_style = template.styles["Heading 1"]
    assert isinstance(heading_style, ParagraphStyle)
    heading_style.font.all_caps = True
    heading_style.font.small_caps = True
    character_style = template.styles.add_style("Authored Character", WD_STYLE_TYPE.CHARACTER)
    assert isinstance(character_style, CharacterStyle)
    character_style.font.name = "Authored Font"
    character_style.font.size = Pt(18)
    character_style.font.small_caps = True
    character_style.font.color.rgb = RGBColor.from_string("135724")
    anchor = template.add_paragraph("{{body}}", style=body_style)
    anchor.runs[0].style = character_style
    anchor.runs[0].font.italic = True
    _add(anchor.runs[0]._r.get_or_add_rPr(), "lang", val="ja-JP")
    paragraph = "Plain **strong** [link](https://example.com) `code`."
    output = _convert(tmp_path, template, f"# Heading:\n{paragraph}\n\n{paragraph}\n")
    for body in output.paragraphs:
        if "Plain " not in body.text:
            continue
        runs = _runs(body)
        for text in ("Plain ", "strong"):
            run = next(run for run in runs if run.text == text)
            assert run.style.style_id == "AuthoredCharacter"
            assert run.italic is True
            assert run.font.name is None and run.font.size is None
            assert _run_property(run, "lang").get(qn("w:val")) == "ja-JP"
            if body.style is not None and body.style.style_id == "Heading1":
                assert run.font.all_caps is False and run.font.small_caps is True
        assert next(run for run in runs if run.text == "strong").bold is True
        assert next(run for run in runs if run.text == "link").style.style_id == "Hyperlink"
        assert next(run for run in runs if run.text == "code").style.style_id == "DocWenInlineCode"
    saved_style = output.styles["Authored Character"]
    assert isinstance(saved_style, CharacterStyle)
    assert saved_style.font.name == "Authored Font" and saved_style.font.size == Pt(18)
    assert saved_style.font.small_caps is True


def test_local_namespace_context_survives_direct_property_projection(tmp_path: Path) -> None:
    document = Document()
    anchor = document.add_paragraph("{{body}}")
    word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    mc = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    anchor._p.insert(
        0,
        parse_xml(
            f'<w:pPr xmlns:w="{word}" xmlns:mc="{mc}" xmlns:future="urn:paragraph-extension" '
            'mc:Ignorable="future"><future:format mode="paragraph"/>'
            '<w:rPr xmlns:mark="urn:mark-extension" mc:Ignorable="mark"><mark:format mode="mark"/>'
            "</w:rPr></w:pPr>"
        ),
    )
    anchor.runs[0]._r.insert(
        0,
        parse_xml(
            f'<w:rPr xmlns:w="{word}" xmlns:mc="{mc}" xmlns:future="urn:character-extension" '
            'mc:Ignorable="future"><future:format mode="character"/></w:rPr>'
        ),
    )
    paragraph_format = extract_body_paragraph_format(document)
    run_format = extract_body_run_format(document)
    assert paragraph_format is not None and run_format is not None
    target = document.add_paragraph("Generated", style="Heading 2")
    original_p, original_r = target._p, target.runs[0]._r
    target_p_pr = target._p.get_or_add_pPr()
    target_r_pr = target.runs[0]._r.get_or_add_rPr()
    _add(target_p_pr, "rPr")  # Exercise the existing paragraph-mark merge.
    paragraph_format.apply_to(target)
    run_format.apply_to(target)
    assert target._p is original_p and target.runs[0]._r is original_r
    assert target._p.pPr is target_p_pr and target.runs[0]._r.rPr is target_r_pr
    path = tmp_path / "namespace-context.docx"
    document.save(str(path))
    output = Document(str(path))
    target = output.paragraphs[1]
    for properties, namespace, mode in (
        (target._p.pPr, "urn:paragraph-extension", "paragraph"),
        (_required(target._p.pPr, qn("w:rPr")), "urn:mark-extension", "mark"),
        (target.runs[0]._r.rPr, "urn:character-extension", "character"),
    ):
        assert properties is not None
        prefix = properties.get(f"{{{mc}}}Ignorable")
        assert prefix is not None and properties.nsmap.get(prefix) == namespace
        extension = _required(properties, f"{{{namespace}}}format")
        assert extension.get("mode") == mode
