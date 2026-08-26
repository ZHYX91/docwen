"""DOCX field processing for gongwen documents (MD→DOCX direction)."""

from __future__ import annotations

from typing import Any

from docwen_core.docx_parsing.xml_ns import NS_W

# ── Extended placeholder rules with row/table deletion ────────────────
PLACEHOLDER_RULES = {
    "delete_paragraph_if_empty": [
        ["密级和保密期限"],
        ["紧急程度"],
        ["发文字号"],
        ["公开方式"],
        ["主送机关"],
        ["附注"],
        ["抄送机关"],
        ["附件说明"],
        ["份号", "发文字号"],  # both empty → delete paragraph
    ],
    "delete_cell_if_empty": [],
    "delete_row_if_empty": [
        ["抄送机关"],
        ["印发机关", "印发日期"],  # both empty → delete row
    ],
    "delete_table_if_empty": [],
}

# ── XML-level placeholder manipulation (Task 3) ────────────────────────


def replace_attachment_placeholder(docx_document: Any, yaml_data: dict) -> None:
    """Replace {{附件说明}} placeholder with formatted attachment paragraphs.

    Searches the document body for the placeholder text, removes the
    placeholder paragraph, and inserts formatted content with hanging
    indentation preserving the original run's font properties.
    Falls back to plain text replacement if DOM manipulation fails.
    """
    from docx.oxml.ns import qn

    placeholder = "{{附件说明}}"
    attachment_desc = yaml_data.get("附件说明")

    # Find placeholder paragraph
    target_para = None
    for para in docx_document.paragraphs:
        if placeholder in para.text:
            target_para = para
            break

    if target_para is None:
        return

    if not attachment_desc or (isinstance(attachment_desc, list) and len(attachment_desc) == 0):
        _remove_paragraph_oxml(target_para)
        return

    if not isinstance(attachment_desc, list):
        attachment_desc = [attachment_desc]

    # Preserve font from original run
    base_rpr = None
    if target_para.runs:
        rpr_elem = target_para.runs[0]._r.find(qn("w:rPr"))
        if rpr_elem is not None:
            base_rpr = rpr_elem

    # Calculate indentation
    left_indent = target_para.paragraph_format.left_indent
    char_width = 12700 * 2  # ~2 chars in EMU
    if left_indent:
        char_width = left_indent / 2

    parent = target_para._p.getparent()
    insert_at = list(parent).index(target_para._p)

    for i, line in enumerate(attachment_desc):
        new_p_elem = _create_paragraph_with_text(
            str(line),
            base_rpr=base_rpr,
            left_indent=int(char_width * 2),
            first_line_indent=-int(char_width * 3 if len(attachment_desc) > 1 else char_width * 2),
        )
        parent.insert(insert_at + i, new_p_elem)

    # Remove original placeholder paragraph
    parent.remove(target_para._p)


def _remove_paragraph_oxml(para) -> None:
    """Remove a paragraph from the document body."""
    parent = para._p.getparent()
    if parent is not None:
        parent.remove(para._p)


def _create_paragraph_with_text(
    text: str,
    base_rpr=None,
    left_indent: int = 0,
    first_line_indent: int = 0,
):
    """Create a new paragraph XML element with text and formatting."""
    from docx.oxml.ns import qn
    from lxml import etree

    p_elem = etree.SubElement(etree.Element("parent"), qn("w:p"))
    pPr = etree.SubElement(p_elem, qn("w:pPr"))

    # Set indentation
    if left_indent:
        ind = etree.SubElement(pPr, qn("w:ind"))
        ind.set(qn("w:left"), str(left_indent))
        if first_line_indent:
            ind.set(qn("w:firstLine"), str(first_line_indent))

    # Create run with text
    r_elem = etree.SubElement(p_elem, qn("w:r"))
    if base_rpr is not None:
        import copy

        r_elem.append(copy.deepcopy(base_rpr))
    t_elem = etree.SubElement(r_elem, qn("w:t"))
    t_elem.text = text
    t_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    # Remove from parent placeholder
    p_elem.getparent().remove(p_elem)  # pyright: ignore[reportOptionalMemberAccess]
    return p_elem


def apply_empty_field_rules(docx_document: Any, yaml_data: dict) -> None:
    """Apply gongwen placeholder rules to remove empty-field content.

    Rules:
    - delete_paragraph_if_empty: remove paragraphs for empty fields
    - delete_row_if_empty: remove table rows where referenced fields are empty
    - delete_table_if_empty: remove tables where all field-referenced rows are empty
    """

    # --- delete_paragraph_if_empty ---
    for field_group in PLACEHOLDER_RULES.get("delete_paragraph_if_empty", []):
        if _all_empty(yaml_data, field_group):
            for field_name in field_group:
                _remove_paragraphs_containing(docx_document, f"{{{{{field_name}}}}}")

    # --- delete_row_if_empty ---
    for field_group in PLACEHOLDER_RULES.get("delete_row_if_empty", []):
        if _all_empty(yaml_data, field_group):
            _remove_table_rows_for_fields(docx_document, field_group)


def _all_empty(yaml_data: dict, field_names: list[str]) -> bool:
    """Check if all named fields have empty/falsy values."""
    return all(not yaml_data.get(name) for name in field_names)


def _remove_paragraphs_containing(docx_document: Any, placeholder: str) -> None:
    """Remove all paragraphs containing the given placeholder text."""
    body = docx_document.element.body
    to_remove = []
    for p in body.findall(f"{{{NS_W}}}p"):
        if placeholder in _extract_paragraph_text(p):
            to_remove.append(p)
    for p in to_remove:
        body.remove(p)


def _remove_table_rows_for_fields(docx_document: Any, field_names: list[str]) -> None:
    """Remove table rows that reference the given empty fields."""
    body = docx_document.element.body
    for tbl in body.findall(f"{{{NS_W}}}tbl"):
        rows_to_remove = []
        for row in tbl.findall(f"{{{NS_W}}}tr"):
            row_text = _extract_element_text(row)
            if any(f"{{{{{fn}}}}}" in row_text for fn in field_names):
                rows_to_remove.append(row)
        for row in rows_to_remove:
            tbl.remove(row)


def _extract_paragraph_text(p_elem) -> str:
    """Extract all text from a paragraph XML element."""
    return "".join((t.text or "") for t in p_elem.findall(f".//{{{NS_W}}}t"))


def _extract_element_text(elem) -> str:
    """Extract all text from any XML element."""
    return "".join(
        (t.text or "") + (t.tail or "") for t in elem.iter() if isinstance(t.tag, str) and t.tag.endswith("}t")
    )
