"""Gongwen-specific DOCX table recognition and plain Markdown projection."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from docwen_core.docx_parsing.table_extraction import (
    markdown_table_lines,
    render_docx_table_rows,
)


@dataclass(frozen=True)
class ExtractedTableParagraph:
    text: str
    anchor_index: int | None
    source: str
    table_index: int
    row_index: int
    cell_index: int
    table_cell_context: str
    table_markdown: str = ""
    is_table_anchor: bool = False
    table_fidelity_risks: tuple[str, ...] = ()


def extract_table_paragraphs(
    doc: Any,
    *,
    table_merge_strategy: str = "fill",
) -> list[ExtractedTableParagraph]:
    """Expose non-empty table cells to Gongwen recognition in source order."""

    items: list[ExtractedTableParagraph] = []
    for table_index, table in enumerate(doc.tables):
        anchor_index = _table_body_index(table, doc)
        rendered = render_docx_table_rows(
            table._tbl,
            cell_text_resolver=lambda cell, _row, _col: _plain_table_cell_text(cell),
            strategy=table_merge_strategy,
        )
        table_markdown = "\n".join(markdown_table_lines(rendered))
        fidelity_risks = _plain_table_fidelity_risks(table._tbl)
        emitted_anchor = False
        seen_cell_elements: set[object] = set()

        for row_index, row in enumerate(table.rows):
            context = "header" if row_index == 0 else "body"
            for cell_index, cell in enumerate(row.cells):
                cell_element = cell._tc
                if cell_element in seen_cell_elements:
                    continue
                seen_cell_elements.add(cell_element)
                text = "<br>".join(
                    paragraph.text.strip() for paragraph in cell.paragraphs if paragraph.text.strip()
                ).replace("|", "\\|")
                if not text:
                    continue

                items.append(
                    ExtractedTableParagraph(
                        text=text,
                        anchor_index=anchor_index,
                        source="table",
                        table_index=table_index,
                        row_index=row_index,
                        cell_index=cell_index,
                        table_cell_context=context,
                        table_markdown=table_markdown if not emitted_anchor else "",
                        is_table_anchor=not emitted_anchor,
                        table_fidelity_risks=fidelity_risks,
                    )
                )
                emitted_anchor = True
        if not emitted_anchor and fidelity_risks:
            # Preserve an evidence anchor even when the plain resolver cannot
            # produce text (for example an image-only or nested-only table).
            # The shared warning layer can then report the loss explicitly.
            items.append(
                ExtractedTableParagraph(
                    text="",
                    anchor_index=anchor_index,
                    source="table",
                    table_index=table_index,
                    row_index=-1,
                    cell_index=-1,
                    table_cell_context="body",
                    is_table_anchor=True,
                    table_fidelity_risks=fidelity_risks,
                )
            )
    return items


def _plain_table_cell_text(cell: Any) -> str:
    """Read direct cell paragraphs without importing Gongwen logic into Core."""

    from docx.oxml.ns import qn

    paragraphs: list[str] = []
    for paragraph in cell.findall(qn("w:p")):
        parts: list[str] = []
        for node in paragraph.iter():
            if node.tag == qn("w:t") and node.text:
                parts.append(node.text)
            elif node.tag == qn("w:tab"):
                parts.append(" ")
            elif node.tag == qn("w:br"):
                parts.append("<br>")
        text = "".join(parts).strip().replace("|", "\\|")
        if text:
            paragraphs.append(text)
    return "<br>".join(paragraphs)


def _plain_table_fidelity_risks(tbl_element: Any) -> tuple[str, ...]:
    """Report cell content deliberately omitted by Gongwen's plain resolver."""

    from docx.oxml.ns import qn

    risks: list[str] = []
    if tbl_element.findall(f".//{qn('w:tc')}/{qn('w:tbl')}"):
        risks.append("nested_table")
    if any(len(properties) > 0 for properties in tbl_element.findall(f".//{qn('w:rPr')}")):
        risks.append("rich_text")
    non_text_tags = (qn("w:drawing"), qn("w:pict"), qn("m:oMath"), qn("m:oMathPara"))
    if any(tbl_element.findall(f".//{tag}") for tag in non_text_tags):
        risks.append("non_text_content")
    unsupported_grid_tags = (qn("w:gridBefore"), qn("w:gridAfter"), qn("w:hMerge"))
    if any(tbl_element.findall(f".//{tag}") for tag in unsupported_grid_tags):
        risks.append("unsupported_grid_layout")
    return tuple(risks)


def _table_body_index(table: Any, doc: Any) -> int | None:
    """Find the owning table's top-level index in ``document.xml``."""

    for index, child in enumerate(doc.element.body):
        if child is table._tbl:
            return index
    return None
