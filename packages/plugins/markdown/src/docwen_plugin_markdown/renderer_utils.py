"""OOXML-level helper functions for MD→DOCX rendering.

Provides:
- OOXML element insertion (soft breaks, shading, font settings)
- Style ID resolution
- Template semantic-style resolution without localized names
- Indentation helpers (quote, list, table)
"""

from __future__ import annotations

import logging

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

logger = logging.getLogger(__name__)

# ── Namespace constants ─────────────────────────────────────────────────
_WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def add_word_soft_break(parent) -> None:
    """Insert a Word ``<w:br>`` soft-break element into a paragraph.

    Appends a new ``<w:r>`` containing ``<w:br/>`` to the paragraph
    element's XML.

    Args:
        parent: A python-docx ``Paragraph`` object or raw ``<w:p>`` element.
    """
    run_elem = OxmlElement("w:r")
    br_elem = OxmlElement("w:br")
    run_elem.append(br_elem)
    if hasattr(parent, "_p"):
        parent._p.append(run_elem)
    else:
        parent.append(run_elem)


def apply_paragraph_shading(paragraph, fill_color: str) -> None:
    """Apply paragraph-level background shading via ``<w:shd>``.

    Args:
        paragraph: A python-docx ``Paragraph`` object.
        fill_color: RRGGBB hex color string (e.g. ``"E7E6E6"``).
    """
    pPr = paragraph._p.get_or_add_pPr()
    existing = pPr.findall(qn("w:shd"))
    for e in existing:
        pPr.remove(e)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_color)
    pPr.append(shd)


def apply_run_shading(run, fill_color: str) -> None:
    """Apply character-level background shading to a run.

    Args:
        run: A python-docx ``Run`` object.
        fill_color: RRGGBB hex color string.
    """
    rPr = run._r.get_or_add_rPr()
    existing = rPr.findall(qn("w:shd"))
    for e in existing:
        rPr.remove(e)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_color)
    rPr.append(shd)


def apply_run_east_asian_font(run, font_name: str) -> None:
    """Set the East Asian font on a run via ``<w:rFonts>``.

    Args:
        run: A python-docx ``Run`` object.
        font_name: East Asian font name (e.g. ``"宋体"``).
    """
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = rPr.makeelement(qn("w:rFonts"), {})
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:eastAsia"), font_name)


def resolve_style_id_by_name(doc, style_name: str) -> str | None:
    """Find the styleId whose ``w:name`` equals *style_name* (case-insensitive).

    Searches all styles in *doc* and returns the matching ``styleId``.

    Args:
        doc: python-docx ``Document``.
        style_name: Display name of the style (e.g. ``"footnote text"``).

    Returns:
        The ``styleId`` string, or ``None`` if not found.
    """
    target = style_name.lower()
    for style in doc.styles:
        style_elem = style._element
        name_elem = style_elem.find(f"{{{_WORD_NS}}}name")
        if name_elem is None:
            continue
        val = name_elem.get(f"{{{_WORD_NS}}}val", "")
        if val.lower() == target:
            sid = style_elem.get(f"{{{_WORD_NS}}}styleId")
            if sid:
                return sid
    return None


def resolve_quote_style(doc, level: int):
    """Return the template-owned quote paragraph style for *level*.

    The distributed templates localize quote display names, so matching names
    inside the plugin would duplicate the i18n table.  Instead, match the stable
    visual contract directly: ``F5F5F5`` shading, a left border, and a 480-twip
    base indent increasing by 240 twips per level.  Every bundled locale and
    compatible custom template can then use its own existing style definition.

    Args:
        doc: A python-docx ``Document`` object.
        level: One-based quote level, from 1 through 9.

    Returns:
        A python-docx paragraph style, or ``None`` if the template has no
        compatible quote style for the requested level.
    """
    if not 1 <= level <= 9:
        return None
    expected_left = 480 + (level - 1) * 240
    for style in doc.styles:
        style_element = style._element
        if style_element.get(qn("w:type")) != "paragraph":
            continue
        p_pr = style_element.find(qn("w:pPr"))
        if p_pr is None:
            continue
        ind = p_pr.find(qn("w:ind"))
        shading = p_pr.find(qn("w:shd"))
        borders = p_pr.find(qn("w:pBdr"))
        left_border = borders.find(qn("w:left")) if borders is not None else None
        if ind is None or shading is None or left_border is None:
            continue
        if ind.get(qn("w:left")) != str(expected_left):
            continue
        if ind.get(qn("w:right")) != "480":
            continue
        if shading.get(qn("w:fill"), "").upper() != "F5F5F5":
            continue
        if left_border.get(qn("w:val")) != "single":
            continue
        if left_border.get(qn("w:sz")) != "24":
            continue
        return style
    return None


def _style_parts(style, expected_type: str):
    """Return ``(style element, pPr, rPr)`` for a requested OOXML type."""
    element = style._element
    if element.get(qn("w:type")) != expected_type:
        return None
    return element, element.find(qn("w:pPr")), element.find(qn("w:rPr"))


def _attr(element, name: str, default: str = "") -> str:
    return element.get(qn(f"w:{name}"), default) if element is not None else default


def resolve_code_block_style(doc):
    """Return the template-owned code-block paragraph style, if present."""
    for style in doc.styles:
        parts = _style_parts(style, "paragraph")
        if parts is None:
            continue
        _element, p_pr, r_pr = parts
        if p_pr is None or r_pr is None:
            continue
        shading = p_pr.find(qn("w:shd"))
        spacing = p_pr.find(qn("w:spacing"))
        indent = p_pr.find(qn("w:ind"))
        fonts = r_pr.find(qn("w:rFonts"))
        size = r_pr.find(qn("w:sz"))
        if (
            _attr(shading, "fill").upper() == "F5F5F5"
            and _attr(spacing, "before") == "120"
            and _attr(spacing, "after") == "120"
            and _attr(spacing, "line") == "240"
            and _attr(indent, "firstLine") == "0"
            and _attr(fonts, "ascii").lower() == "consolas"
            and _attr(size, "val") == "20"
        ):
            return style
    return None


def resolve_inline_code_style(doc):
    """Return the template-owned inline-code character style, if present."""
    for style in doc.styles:
        parts = _style_parts(style, "character")
        if parts is None:
            continue
        _element, _p_pr, r_pr = parts
        if r_pr is None:
            continue
        fonts = r_pr.find(qn("w:rFonts"))
        shading = r_pr.find(qn("w:shd"))
        if _attr(fonts, "ascii").lower() == "consolas" and _attr(shading, "fill").upper() == "F0F0F0":
            return style
    return None


def resolve_formula_block_style(doc):
    """Return the template-owned block-formula paragraph style, if present."""
    for style in doc.styles:
        parts = _style_parts(style, "paragraph")
        if parts is None:
            continue
        _element, p_pr, _r_pr = parts
        if p_pr is None:
            continue
        spacing = p_pr.find(qn("w:spacing"))
        indent = p_pr.find(qn("w:ind"))
        alignment = p_pr.find(qn("w:jc"))
        if (
            _attr(spacing, "before") == "120"
            and _attr(spacing, "after") == "120"
            and not _attr(spacing, "line")
            and _attr(indent, "firstLine") == "0"
            and _attr(alignment, "val") == "center"
            and p_pr.find(qn("w:shd")) is None
            and p_pr.find(qn("w:pBdr")) is None
        ):
            return style
    return None


def resolve_inline_formula_style(doc):
    """Return the template-owned inline-formula character style, if present."""
    for style in doc.styles:
        parts = _style_parts(style, "character")
        if parts is None:
            continue
        element, _p_pr, r_pr = parts
        if element.get(qn("w:customStyle")) != "1":
            continue
        priority = element.find(qn("w:uiPriority"))
        if (
            element.find(qn("w:qFormat")) is not None
            and _attr(priority, "val") == "29"
            and (r_pr is None or len(r_pr) == 0)
        ):
            return style
    return None


def resolve_list_block_style(doc):
    """Return the template-owned list paragraph style, if present."""
    for style in doc.styles:
        parts = _style_parts(style, "paragraph")
        if parts is None:
            continue
        _element, p_pr, _r_pr = parts
        if p_pr is None or p_pr.find(qn("w:contextualSpacing")) is None:
            continue
        indent = p_pr.find(qn("w:ind"))
        if _attr(indent, "left") == "720" and _attr(indent, "firstLine") == "0":
            return style
    return None


def resolve_table_content_style(doc):
    """Return the template-owned table-cell paragraph style, if present."""
    for style in doc.styles:
        parts = _style_parts(style, "paragraph")
        if parts is None:
            continue
        _element, p_pr, r_pr = parts
        if p_pr is None or r_pr is None:
            continue
        spacing = p_pr.find(qn("w:spacing"))
        indent = p_pr.find(qn("w:ind"))
        alignment = p_pr.find(qn("w:jc"))
        size = r_pr.find(qn("w:sz"))
        if (
            _attr(spacing, "before") == "0"
            and _attr(spacing, "after") == "0"
            and _attr(spacing, "line") == "240"
            and _attr(indent, "firstLine") == "0"
            and _attr(alignment, "val") == "center"
            and _attr(size, "val") == "21"
        ):
            return style
    return None


def resolve_table_style(doc, semantic_key: str):
    """Return a localized template table style by its border contract."""
    expected = semantic_key.strip().lower()
    if expected not in {"three_line_table", "table_grid"}:
        return None
    for style in doc.styles:
        parts = _style_parts(style, "table")
        if parts is None:
            continue
        element, _p_pr, _r_pr = parts
        table_pr = element.find(qn("w:tblPr"))
        borders = table_pr.find(qn("w:tblBorders")) if table_pr is not None else None
        if borders is None:
            continue
        edges = {edge.tag.rsplit("}", 1)[-1]: edge for edge in borders}
        if expected == "three_line_table":
            if (
                set(edges) == {"top", "bottom"}
                and all(_attr(edges[name], "val") == "single" for name in edges)
                and all(_attr(edges[name], "sz") == "12" for name in edges)
            ):
                return style
        elif set(edges) == {"top", "left", "bottom", "right", "insideH", "insideV"} and all(
            _attr(edge, "val") == "single" and _attr(edge, "sz") == "4" for edge in edges.values()
        ):
            return style
    return None


def apply_quote_style(paragraph, *, style=None, level: int = 1) -> None:
    """Apply a template quote style or its visible direct-format fallback.

    Args:
        paragraph: A python-docx ``Paragraph`` object.
        style: Template-owned paragraph style resolved by
            :func:`resolve_quote_style`.
        level: One-based quote level used by the fallback formatter.
    """
    if style is not None:
        paragraph.style = style
        return

    level = min(max(level, 1), 9)
    pPr = paragraph._p.get_or_add_pPr()
    for existing in pPr.findall(qn("w:ind")):
        pPr.remove(existing)
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), str(480 + (level - 1) * 240))
    ind.set(qn("w:right"), "480")
    ind.set(qn("w:firstLine"), "0")
    pPr.append(ind)

    spacing = pPr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        pPr.append(spacing)
    spacing.set(qn("w:before"), "120")
    spacing.set(qn("w:after"), "120")

    apply_paragraph_shading(paragraph, "F5F5F5")

    for existing in pPr.findall(qn("w:pBdr")):
        pPr.remove(existing)
    borders = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    color_value = max(0x66, 0xCC - (level - 1) * 0x0B)
    border_color = f"{color_value:02X}{color_value:02X}{color_value:02X}"
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), "24")
    left.set(qn("w:space"), "12")
    left.set(qn("w:color"), border_color)
    borders.append(left)
    pPr.append(borders)


def apply_list_indent(paragraph, level: int = 0) -> None:
    """Apply nested list contextual indent to a paragraph.

    Uses ``w:ind`` with left offset based on nesting depth.

    Args:
        paragraph: A python-docx ``Paragraph`` object.
        level: Zero-based nesting depth (0 = top-level list).
    """
    indent_twips = 360 * (level + 1)
    pPr = paragraph._p.get_or_add_pPr()
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), str(indent_twips))
    ind.set(qn("w:hanging"), "360")
    pPr.append(ind)


def set_table_left_indent(table, indent_twips: int) -> None:
    """Set the ``<w:tblInd>`` left indent on a Word table.

    Args:
        table: A python-docx ``Table`` object.
        indent_twips: Left indent amount in twips.
    """
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    existing = tblPr.findall(qn("w:tblInd"))
    for e in existing:
        tblPr.remove(e)
    tbl_ind = OxmlElement("w:tblInd")
    tbl_ind.set(qn("w:w"), str(indent_twips))
    tbl_ind.set(qn("w:type"), "dxa")
    tblPr.append(tbl_ind)


def enable_table_header_row_formatting(table) -> None:
    """Enable first-row conditional formatting via ``<w:tblLook>``.

    Tells Word to apply the table style's header-row banding to the
    first row(s).
    """
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    existing = tblPr.findall(qn("w:tblLook"))
    for e in existing:
        tblPr.remove(e)
    tblLook = OxmlElement("w:tblLook")
    tblLook.set(qn("w:firstRow"), "1")
    tblLook.set(qn("w:lastRow"), "0")
    tblLook.set(qn("w:firstColumn"), "0")
    tblLook.set(qn("w:lastColumn"), "0")
    tblLook.set(qn("w:noHBand"), "0")
    tblLook.set(qn("w:noVBand"), "1")
    tblPr.append(tblLook)


def apply_three_line_table_borders(table) -> None:
    """Apply academic three-line table borders directly to a Word table.

    The renderer expresses the built-in three-line-table semantics locally:
    thick top and bottom table borders, with the header separator handled at
    cell level for Word/WPS compatibility.
    """
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    existing = tblPr.findall(qn("w:tblBorders"))
    for e in existing:
        tblPr.remove(e)
    tbl_borders = OxmlElement("w:tblBorders")
    for edge in ("top", "bottom"):
        border = OxmlElement(f"w:{edge}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "12")
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), "auto")
        tbl_borders.append(border)
    tblPr.append(tbl_borders)


def apply_table_grid_borders(table) -> None:
    """Apply Table Grid-like borders directly to a Word table."""
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    existing = tblPr.findall(qn("w:tblBorders"))
    for e in existing:
        tblPr.remove(e)
    tbl_borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement(f"w:{edge}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "4")
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), "auto")
        tbl_borders.append(border)
    tblPr.append(tbl_borders)


def apply_header_row_bottom_border(row) -> None:
    """Apply the three-line-table header separator to every cell in *row*."""
    for cell in row.cells:
        tc_pr = cell._tc.get_or_add_tcPr()
        tc_borders = tc_pr.find(qn("w:tcBorders"))
        if tc_borders is None:
            tc_borders = OxmlElement("w:tcBorders")
            tc_pr.append(tc_borders)

        existing = tc_borders.find(qn("w:bottom"))
        if existing is not None:
            continue

        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "4")
        bottom.set(qn("w:space"), "0")
        bottom.set(qn("w:color"), "000000")
        tc_borders.append(bottom)
