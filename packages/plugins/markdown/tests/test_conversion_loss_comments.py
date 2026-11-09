"""Conversion losses remain visible and diagnosed independently of comments."""

from __future__ import annotations

from io import BytesIO

import pytest
from docx import Document

from docwen_core.conversion_loss import ConversionLoss, LossCollector
from docwen_core.docx_comments import append_docx_comment
from docwen_plugin_markdown.mistune_extensions import parse_markdown_text
from docwen_plugin_markdown.renderer import MdToDocxRenderer
from docwen_plugin_markdown.to_docx.notes import extract_notes_from_ast, normalize_note_syntax

pytestmark = pytest.mark.contract


def test_loss_collector_deduplicates_occurrences_not_originals():
    collector = LossCollector()
    one = ConversionLoss("CONVERSION-LOSS-TEST", "unsupported", "placeholder", "line 2", "paragraph 1", "$x$")
    two = ConversionLoss("CONVERSION-LOSS-TEST", "unsupported", "placeholder", "line 3", "paragraph 2", "$x$")
    assert collector.add(one)
    assert not collector.add(one)
    assert collector.add(two)
    assert len(collector.diagnostics) == 2
    assert collector.diagnostics[0].location == "paragraph 1"
    assert "line 2" in collector.diagnostics[0].message


@pytest.mark.parametrize("comments", [True, False])
def test_formula_loss_preserves_template_comment_and_diagnostics(comments):
    document = Document()
    original_run = document.add_paragraph("Existing").runs[0]
    append_docx_comment(document, [original_run], text="User review", author="User")
    renderer = MdToDocxRenderer(document, loss_comments=comments)
    renderer.render(parse_markdown_text(r"Before $x\hspace{1em}y$ after"))
    [loss] = renderer.losses.records
    assert loss.original == r"$x\hspace{1em}y$"
    assert renderer.loss_comments_added == int(comments)
    saved = BytesIO()
    document.save(saved)
    restored = Document(saved)
    assert len(restored.comments) == 1 + int(comments)
    assert next(iter(restored.comments)).text == "User review"
    assert next(iter(restored.comments)).author == "User"
    assert restored.paragraphs[1].text == "Before [Formula could not be converted] after"


def test_comment_failure_retains_loss_and_failure_diagnostics(monkeypatch):
    import docwen_plugin_markdown.renderer as renderer_module

    def fail(*args, **kwargs):
        raise OSError("cannot write")

    monkeypatch.setattr(renderer_module, "append_docx_comment", fail)
    renderer = MdToDocxRenderer(Document())
    renderer.render(parse_markdown_text(r"$x\hspace{1em}y$"))
    assert len(renderer.losses.records) == 1
    assert renderer.warnings[0][0] == "CONVERSION-LOSS-COMMENT-FAILED"
    assert renderer.loss_comments_added == 0


@pytest.mark.parametrize("kind", ["footnote", "endnote"])
@pytest.mark.parametrize("comments", [True, False])
def test_note_formula_loss_anchors_only_first_body_reference(kind, comments):
    source = f"First[^{kind}:a] again[^{kind}:a].\n\n[^{kind}:a]: $x^2$\n"
    ast, notes = extract_notes_from_ast(parse_markdown_text(normalize_note_syntax(source)))
    document = Document()
    renderer = MdToDocxRenderer(document, note_ctx=notes, loss_comments=comments)
    renderer.render(ast)
    [loss] = renderer.losses.records
    assert loss.code == "CONVERSION-LOSS-NOTE-FORMULA"
    assert loss.original == "$x^2$"
    assert kind in loss.output_location
    assert renderer.loss_comments_added == int(comments)
    assert len(document._element.xpath(f".//w:{kind}Reference")) == 2
    assert len(document._element.xpath(".//w:commentRangeStart")) == int(comments)
    elements = notes.footnote_elements if kind == "footnote" else notes.endnote_elements
    assert "$x^2$" in "".join(str(text) for text in elements[0].itertext())


def test_successful_formula_has_no_loss_or_comments():
    document = Document()
    renderer = MdToDocxRenderer(document)
    renderer.render(parse_markdown_text("$x^2$"))
    assert not renderer.losses.records
    assert not document._element.xpath(".//w:commentRangeStart")


@pytest.mark.parametrize("token", ["mi", "mo", "mtext", "mrow"])
def test_unsupported_mathvariant_cannot_silently_convert(token):
    from docwen_core.formula import mathml_to_omml

    body = "<mi>x</mi>" if token == "mrow" else "x"
    assert mathml_to_omml(f'<math><{token} mathvariant="double-struck">{body}</{token}></math>') is None


def test_summary_reports_annotations_only_when_all_succeeded():
    collector = LossCollector()
    assert collector.summary() is None
    collector.add(ConversionLoss("CONVERSION-LOSS-TEST", "lost", "placeholder"))
    unannotated = collector.summary()
    annotated = collector.summary(annotated_count=1)
    assert unannotated is not None and annotated is not None
    assert "Comments were added" not in unannotated.message
    assert "Comments were added" in annotated.message


def test_loss_messages_localize_placeholder_and_comment():
    document = Document()
    renderer = MdToDocxRenderer(
        document,
        loss_messages={
            "conversion_loss.formula_unconverted": "〔公式未能转换〕",
            "conversion_loss.formula_reason": "此公式未能保留。",
            "conversion_loss.formula_change": "已插入提示，请核对原式。",
            "conversion_loss.original": "原式：{original}",
        },
    )
    renderer.render(parse_markdown_text(r"$x\hspace{1em}y$"))
    assert "〔公式未能转换〕" in document.paragraphs[0].text
    assert "原式：" in next(iter(document.comments)).text
    assert "此公式未能保留。" in next(iter(document.comments)).text


@pytest.mark.parametrize("comments", [True, False])
def test_deferred_loss_locations_follow_final_template_relocation(comments):
    document = Document()
    document.add_paragraph("Template heading")
    placeholder = document.add_paragraph("Body placeholder")
    document.add_paragraph("Template closing")
    renderer = MdToDocxRenderer(document, loss_comments=comments, defer_loss_comments=True)
    [rendered] = renderer.render(parse_markdown_text(r"Before $x\hspace{1em}y$ after"))
    [before] = renderer.losses.records
    assert before.output_location.startswith("Paragraph 4,")
    assert renderer.loss_comments_added == 0
    placeholder._p.addprevious(rendered._p)
    owner = placeholder._p.getparent()
    assert owner is not None
    owner.remove(placeholder._p)
    renderer.finalize_loss_comments()
    renderer.finalize_loss_comments()
    [after] = renderer.losses.records
    assert after.output_location.startswith("Paragraph 2,")
    assert renderer.loss_comments_added == int(comments)
    saved = BytesIO()
    document.save(saved)
    reopened = Document(saved)
    assert "Formula could not be converted" in reopened.paragraphs[1].text
    assert len(reopened.comments) == int(comments)
    if comments:
        assert "Paragraph 2," in next(iter(reopened.comments)).text
        assert "Paragraph 4," not in next(iter(reopened.comments)).text


@pytest.mark.parametrize("comments", [True, False])
def test_real_converter_loss_coordinates_match_saved_template_position(tmp_path, comments):
    from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter

    from .conftest import make_context

    source = tmp_path / "loss.md"
    source.write_text(r"Before $x\hspace{1em}y$ after", encoding="utf-8")
    template_path = tmp_path / "template.docx"
    template = Document()
    template.add_paragraph("Template heading")
    template.add_paragraph("{{body}}")
    template.add_paragraph("Template closing")
    template.save(template_path)
    context, _workspace = make_context(
        str(source), options={"template_name": str(template_path)}, config_values={"conversion.loss_comments": comments}
    )
    result = MdToDocxConverter().convert(context)
    assert result.success, result.error
    [loss] = [item for item in result.diagnostics if item.code == "CONVERSION-LOSS-FORMULA"]
    assert loss.location.startswith("Paragraph 2,")
    reopened = Document(result.artifacts[0].staging_path)
    assert "Formula could not be converted" in reopened.paragraphs[1].text
    assert len(reopened.comments) == int(comments)
    if comments:
        assert "Paragraph 2," in next(iter(reopened.comments)).text
