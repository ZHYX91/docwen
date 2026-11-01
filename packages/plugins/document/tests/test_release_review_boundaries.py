"""Source-to-artifact order and literal prose oracles for the release review."""

from __future__ import annotations

import hashlib
from pathlib import Path

import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from markdown_it import MarkdownIt
from PIL import Image

from docwen_core.docx_parsing.format_features import DocxMarkdownSyntaxConfig
from docwen_plugin_document.shared.markdown_runs import render_paragraph_runs

from ._to_markdown_standard_parity_support import _convert_document_fixture_to_markdown

pytestmark = pytest.mark.contract


def _inline(markdown: str):
    return [child for token in MarkdownIt().parse(markdown) if token.type == "inline" for child in token.children or []]


@pytest.mark.parametrize("split_runs", [False, True])
def test_release_review_literal_prose_has_only_text_ast(split_runs: bool) -> None:
    source = r"# *服务* _name_ [click](https://invalid.test) &amp; <b>raw</b> \path `tick` ![image](x)"
    paragraph = Document().add_paragraph()
    for fragment in source if split_runs else [source]:
        paragraph.add_run(fragment)
    rendered = render_paragraph_runs(paragraph, syntax_config=DocxMarkdownSyntaxConfig())
    children = _inline(rendered)
    assert {child.type for child in children} == {"text"}
    assert "".join(child.content for child in children) == source


@pytest.mark.parametrize("source", ["- plain", "+ plain", "1. plain", "2) plain", "===", "---"])
def test_release_review_literal_block_prefix(source: str) -> None:
    paragraph = Document().add_paragraph()
    for character in source:
        paragraph.add_run(character)
    rendered = render_paragraph_runs(paragraph, syntax_config=DocxMarkdownSyntaxConfig())
    tokens = MarkdownIt().parse(rendered)
    assert tokens[0].type == "paragraph_open"
    assert "".join(child.content for child in _inline(rendered)) == source


def test_release_review_real_format_and_code_remain_semantic() -> None:
    paragraph = Document().add_paragraph()
    paragraph.add_run("*literal* ")
    paragraph.add_run("bold").bold = True
    paragraph.add_run(" ")
    code = paragraph.add_run("`literal` &amp;")
    properties = code._r.get_or_add_rPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), "D9D9D9")
    properties.append(shading)
    children = _inline(render_paragraph_runs(paragraph, syntax_config=DocxMarkdownSyntaxConfig()))
    assert any(child.type == "strong_open" for child in children)
    assert [child.content for child in children if child.type == "code_inline"] == ["`literal` &amp;"]
    assert "*literal* " in "".join(child.content for child in children if child.type == "text")


@pytest.mark.parametrize("same_run", [False, True])
@pytest.mark.parametrize("container", ["body", "sdt", "sdt-table"])
def test_release_review_docx_text_image_order_and_resource_bytes(
    tmp_path: Path, same_run: bool, container: str
) -> None:
    image_paths = []
    for index, color in enumerate(("red", "blue")):
        path = tmp_path / f"source-{index}.png"
        Image.new("RGB", (4, 4), color).save(path)
        image_paths.append(path)
    document = Document()
    paragraph = document.add_paragraph()
    run = paragraph.add_run("BEFORE")
    run.add_picture(str(image_paths[0]))
    run = run if same_run else paragraph.add_run()
    run.add_text("MIDDLE")
    run.add_picture(str(image_paths[1]))
    run = run if same_run else paragraph.add_run()
    run.add_text("AFTER")
    if container != "body":
        sdt = OxmlElement("w:sdt")
        content = OxmlElement("w:sdtContent")
        if container == "sdt-table":
            table = document.add_table(rows=1, cols=1)
            table.cell(0, 0)._tc.append(paragraph._p)
            content.append(table._tbl)
        else:
            content.append(paragraph._p)
        sdt.append(content)
        document._element.body.insert(0, sdt)
    markdown = _convert_document_fixture_to_markdown(
        tmp_path, document, request_id="ordered", options={"image_link_style": "markdown_embed"}
    )
    children = _inline(markdown.split("---", 2)[-1])
    assert sum(child.type == "image" for child in children) == 2
    if container != "sdt-table":
        assert [child.type for child in children] == ["text", "image", "text", "image", "text"]
        assert [child.content for child in children if child.type == "text"] == ["BEFORE", "MIDDLE", "AFTER"]
    # Inspect independently of exact resource filenames.
    positions = [markdown.index(text) for text in ("BEFORE", "MIDDLE", "AFTER")]
    image_positions = [index for index in range(len(markdown)) if markdown.startswith("![", index)]
    assert len(image_positions) == 2
    assert positions[0] < image_positions[0] < positions[1] < image_positions[1] < positions[2]
    output_images = list((tmp_path / "ordered-staging").glob("*.png"))
    assert sorted(hashlib.sha256(path.read_bytes()).hexdigest() for path in output_images) == sorted(
        hashlib.sha256(path.read_bytes()).hexdigest() for path in image_paths
    )


def test_release_review_reference_only_paragraph_remains_visible() -> None:
    from types import SimpleNamespace

    from docwen_plugin_document.to_markdown.converter import DocxToMarkdownConverter

    paragraph = Document().add_paragraph()
    reference = OxmlElement("w:footnoteReference")
    reference.set(qn("w:id"), "1")
    paragraph.add_run()._r.append(reference)
    converter = DocxToMarkdownConverter()
    converter._note_extractor = SimpleNamespace(get_reference_text=lambda kind, number: "[^1]")
    lines, stats = converter._process_paragraph(paragraph._p, {id(paragraph._p): paragraph})
    assert lines == ["[^1]", ""]
    assert stats["paragraphs"] == 1


def test_release_review_unformatted_heading_is_literal() -> None:
    from docwen_plugin_document.to_markdown.converter import DocxToMarkdownConverter

    paragraph = Document().add_paragraph("*literal* <br>", style="Heading 1")
    converter = DocxToMarkdownConverter()
    converter._preserve_heading_formatting = False
    lines, _stats = converter._process_paragraph(paragraph._p, {id(paragraph._p): paragraph})
    tokens = MarkdownIt().parse("\n".join(lines))
    assert tokens[0].type == "heading_open"
    assert {child.type for child in tokens[1].children or []} == {"text"}
    children = tokens[1].children
    assert children is not None
    assert children[0].content == "*literal* <br>"
