"""Thin adapter: Markdown formula AST nodes → DOCX OMML elements."""

from __future__ import annotations

from collections.abc import Callable, Mapping
from typing import Any

from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from docwen_core.conversion_loss import ConversionLoss, loss_text
from docwen_core.markdown_math import serialize_formula


def _latex_to_omml(latex: str) -> Any:
    """Convert LaTeX string to lxml OMML element via core formula pipeline."""
    import lxml.etree as etree

    from docwen_core.formula import latex_to_mathml, mathml_to_omml

    if not latex.strip():
        raise ValueError("empty formula")
    mathml = latex_to_mathml(latex)
    if not mathml:
        raise ValueError("LaTeX could not be parsed as MathML")
    omml = mathml_to_omml(mathml)
    if not omml:
        raise ValueError("MathML contains an unsupported or invalid mathematical structure")
    return etree.fromstring(omml.encode("utf-8"))


def render_block_formula(
    document,
    latex: str,
    *,
    block_style=None,
    in_list: bool = False,
    source: str | None = None,
    on_loss: Callable[[ConversionLoss, Any], None] | None = None,
    loss_messages: Mapping[str, str] | None = None,
) -> bool:
    """Insert explicit display math; list content and its marker stay left."""
    p = document.add_paragraph(style=block_style)
    if in_list:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    elif block_style is None:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    alignment = p.alignment if p.alignment is not None else p.style.paragraph_format.alignment
    math_alignment = {WD_ALIGN_PARAGRAPH.LEFT: "left", WD_ALIGN_PARAGRAPH.RIGHT: "right"}.get(alignment, "center")
    return append_formula(
        p, latex, block=True, source=source, alignment=math_alignment, on_loss=on_loss, loss_messages=loss_messages
    )


def _apply_character_style(omml: Any, style_reference: str | None) -> None:
    """Apply a Word character style to formula runs and control properties."""
    if not style_reference:
        return

    import lxml.etree as etree

    math_ns = "http://schemas.openxmlformats.org/officeDocument/2006/math"
    word_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    targets = list(omml.iter(f"{{{math_ns}}}r")) + list(omml.iter(f"{{{math_ns}}}ctrlPr"))
    for target in targets:
        r_pr = target.find(qn("w:rPr"))
        if r_pr is None:
            r_pr = etree.Element(f"{{{word_ns}}}rPr")
            # CT_R requires math properties before Word run properties.
            target.insert(1 if target.find(qn("m:rPr")) is not None else 0, r_pr)
        r_style = r_pr.find(qn("w:rStyle"))
        if r_style is None:
            r_style = etree.Element(f"{{{word_ns}}}rStyle")
            r_pr.insert(0, r_style)
        r_style.set(qn("w:val"), style_reference)


def formula_loss_location(paragraph: Any, run: Any, messages: Mapping[str, str] | None) -> str:
    """Use visible document coordinates rather than an internal XML path."""
    root = paragraph._p.getroottree().getroot()
    paragraphs = list(root.iter(qn("w:p")))
    paragraph_number = paragraphs.index(paragraph._p) + 1
    item = list(paragraph._p.iter(qn("w:r"))).index(run._r) + 1
    cell = next((node for node in paragraph._p.iterancestors() if node.tag == qn("w:tc")), None)
    if cell is None:
        return loss_text(messages, "docx_paragraph", "Paragraph {paragraph}, item {item}").format(
            paragraph=paragraph_number, item=item
        )
    row = cell.getparent()
    table = row.getparent()
    return loss_text(
        messages, "docx_cell", "Table {table}, row {row}, column {column}, paragraph {paragraph}, item {item}"
    ).format(
        table=list(root.iter(qn("w:tbl"))).index(table) + 1,
        row=list(table.iterchildren(qn("w:tr"))).index(row) + 1,
        column=list(row.iterchildren(qn("w:tc"))).index(cell) + 1,
        paragraph=list(cell.iter(qn("w:p"))).index(paragraph._p) + 1,
        item=item,
    )


def append_formula(
    paragraph,
    latex: str,
    *,
    block: bool,
    style_reference: str | None = None,
    source: str | None = None,
    alignment: str = "left",
    on_loss: Callable[[ConversionLoss, Any], None] | None = None,
    loss_messages: Mapping[str, str] | None = None,
) -> bool:
    """Write a math node in any paragraph, including lists and table cells.

    The OOXML wrapper carries display identity. Alignment belongs to the
    container and never changes that identity. A failed conversion keeps the
    original formula in a loss comment, with a visible occurrence marker.
    """
    try:
        omml = _latex_to_omml(latex)
    except ValueError:
        original = source if source is not None else serialize_formula(latex, block=block)
        run = paragraph.add_run(loss_text(loss_messages, "formula_unconverted", "[Formula could not be converted]"))
        if style_reference:
            run.style = paragraph.part.document.styles.get_by_id(style_reference, WD_STYLE_TYPE.CHARACTER)
        location = formula_loss_location(paragraph, run, loss_messages)
        loss = ConversionLoss(
            code="CONVERSION-LOSS-FORMULA",
            reason=loss_text(
                loss_messages, "formula_reason", "This formula could not be represented as a Word equation."
            ),
            change=loss_text(
                loss_messages, "formula_change", "A placeholder was inserted; check the original formula."
            ),
            output_location=location,
            original=original,
        )
        if on_loss is not None:
            on_loss(loss, run)
        else:
            from docwen_core.docx_comments import append_docx_comment

            append_docx_comment(paragraph.part.document, [run], text=loss.message(loss_messages))
        return False
    _apply_character_style(omml, style_reference)
    if block:
        display = OxmlElement("m:oMathPara")
        properties = OxmlElement("m:oMathParaPr")
        justification = OxmlElement("m:jc")
        justification.set(qn("m:val"), alignment)
        properties.append(justification)
        display.append(properties)
        display.append(omml)
        paragraph._p.append(display)
    else:
        paragraph._p.append(omml)
    return True
