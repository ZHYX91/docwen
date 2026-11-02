"""Native presentation prose stays literal in independently rendered Markdown."""

from __future__ import annotations

import pytest
from lxml import html
from markdown_it import MarkdownIt
from pptx import Presentation
from pptx.util import Inches

from ._presentation_to_md_support import Path, _inject_pptx_sections, _run_request
from ._presentation_to_md_support import pipeline as pipeline

pytestmark = pytest.mark.contract


def test_saved_markdown_preserves_literal_text_and_only_omits_title_shape(pipeline, tmp_path):
    literal = "<b>literal</b> [manual](https://example.com) *not italic* &copy;"
    presentation = Presentation()
    presentation.core_properties.title = "Guide *literal*"
    slide = presentation.slides.add_slide(presentation.slide_layouts[5])
    assert slide.shapes.title is not None
    slide.shapes.title.text = literal
    body = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(1))
    body.text_frame.text = literal
    block_text = "# literal heading\n- not a bullet\n1. not ordered\n---\n$$price$$"
    block = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(8), Inches(2))
    block.text_frame.text = block_text
    note = "note <b>tag</b> and *stars*"
    notes = slide.notes_slide.notes_text_frame
    assert notes is not None
    notes.text = note
    source = tmp_path / "source.pptx"
    presentation.save(source)
    section_title = "Section <b>literal</b>"
    _inject_pptx_sections(source, {section_title: 0})
    original = source.read_bytes()
    output = tmp_path / "output"
    output.mkdir()
    _plugin, manager, _workspace = pipeline
    result = _run_request(manager, source, "pptx", output, export_notes=True)
    assert result.success, result.error
    content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
    body_markdown = content.split("---", 2)[2]
    rendered = html.fragment_fromstring(
        MarkdownIt("commonmark", {"html": True}).render(body_markdown), create_parent=True
    )
    assert [node.text_content() for node in rendered.xpath("//h1")] == ["Guide *literal*", section_title]
    assert [node.text_content() for node in rendered.xpath("//h2")] == [literal]
    assert rendered.text_content().count(literal) == 2
    assert note in rendered.text_content()
    for line in block_text.splitlines():
        assert line in rendered.text_content()
    assert rendered.xpath("//a | //b | //em | //ul | //ol | //hr") == []
    assert source.read_bytes() == original
