"""Page/section break detection and border group tracking for DOCX->MD."""

from __future__ import annotations

from typing import Any

from docwen_core.docx_parsing.style_properties import paragraph_property_sources
from docwen_core.docx_parsing.xml_ns import NS_W

# ── Break detection ─────────────────────────────────────────────────────


def _is_in_rejected_revision(element: Any, paragraph: Any) -> bool:
    parent = element.getparent() if hasattr(element, "getparent") else None
    while parent is not None and parent is not paragraph:
        local_name = parent.tag.split("}")[-1] if "}" in (parent.tag or "") else (parent.tag or "")
        if local_name in {"del", "moveFrom"}:
            return True
        parent = parent.getparent() if hasattr(parent, "getparent") else None
    return False


def detect_page_break(para: Any) -> bool:
    """Return True if the paragraph contains a page break."""
    nsmap = {"w": NS_W}
    breaks = para._p.findall(".//w:br", nsmap)
    return any(br.get(f"{{{NS_W}}}type") == "page" and not _is_in_rejected_revision(br, para._p) for br in breaks)


def detect_section_break(para: Any) -> tuple[str, str] | None:
    """Read the current paragraph section properties, excluding revision history."""
    elem = para._p
    sect_pr = elem.find(f"{{{NS_W}}}pPr/{{{NS_W}}}sectPr")
    if sect_pr is None:
        return None
    section_type = sect_pr.find(f"{{{NS_W}}}type")
    stype = section_type.get(f"{{{NS_W}}}val", "nextPage") if section_type is not None else "nextPage"
    return (stype, "")


def detect_all_breaks(para: Any) -> list[tuple[str, str]]:
    """Return ordered list of (break_type, details) tuples."""
    breaks: list[tuple[str, str]] = []
    if detect_page_break(para):
        breaks.append(("page", ""))
    sb = detect_section_break(para)
    if sb is not None:
        breaks.append(("section", sb[0]))
    return breaks


# ── Horizontal rule ─────────────────────────────────────────────────────


def detect_horizontal_rule(para: Any) -> bool:
    """Return True if the paragraph is a horizontal rule (border + empty text)."""
    info = extract_paragraph_border_info(para)
    if info.get("bottom"):
        element = getattr(para, "_p", None)
        if element is None:
            return False
        for child in element.iter():
            if _is_in_rejected_revision(child, element):
                continue
            name = child.tag.rsplit("}", 1)[-1]
            if name in {"drawing", "pict", "oMath", "oMathPara"}:
                return False
            if name == "t" and (child.text or "").strip():
                return False
        return True
    return False


def extract_paragraph_border_info(para: Any) -> dict:
    """Resolve each border edge through direct properties and paragraph styles.

    An explicit ``none``/``nil`` edge clears its inherited counterpart; an
    unrelated direct paragraph property must not hide the style's borders.
    """
    info: dict = {}
    seen: set[str] = set()
    for pPr in paragraph_property_sources(para):
        pBdr = pPr.find(f"{{{NS_W}}}pBdr")
        if pBdr is None:
            continue
        for direction in ("top", "bottom", "left", "right", "between"):
            if direction in seen:
                continue
            bdr = pBdr.find(f"{{{NS_W}}}{direction}")
            if bdr is not None:
                seen.add(direction)
                val = bdr.get(f"{{{NS_W}}}val")
                if val and val not in ("none", "nil"):
                    info[direction] = val
    return info


# ── Border group tracker ────────────────────────────────────────────────


class BorderGroupTracker:
    """Track border groups to emit separators between grouped paragraphs.

    Rules (from old break_processor.py):
    - top border starts group → emit separator before paragraph
    - between border inside group → emit separator between paragraphs
    - bottom border ends group → emit separator after paragraph
    - adjacent group boundaries collapse (don't double-emit)
    """

    def __init__(self, separator: str = "---") -> None:
        self.is_in_group = False
        self._just_closed = False
        self._just_opened = False
        self._bottom_only_pending = False
        self._separator = separator

    def process_paragraph(self, border_info: dict | None = None) -> list[str]:
        """Process one paragraph's border info, return separator lines to emit.

        Returns empty list or a list with one separator line.
        """
        result: list[str] = []
        border_info = border_info or {}

        top = border_info.get("top")
        bottom = border_info.get("bottom")
        between = border_info.get("between")

        has_top = top is not None
        has_bottom = bottom is not None
        has_between = between is not None
        has_border = has_top or has_bottom or has_between

        self._just_opened = False
        self._just_closed = False

        # Word's auto-generated horizontal rules are commonly represented by
        # an empty paragraph with only a bottom border.  Keep that border
        # pending until the following non-border paragraph (or ``finalize``),
        # matching the legacy converters' state-machine behavior.
        if not has_border and self.is_in_group and self._bottom_only_pending:
            if self._separator:
                result.append(self._separator)
            self.is_in_group = False
            self._bottom_only_pending = False
            self._just_closed = True
            return result

        if has_top and not self.is_in_group:
            if self._separator and not self._just_closed:
                result.append(self._separator)
            self.is_in_group = True
            self._just_opened = True

        if self._separator and has_between and self.is_in_group:
            result.append(self._separator)

        if self._separator and has_bottom and self.is_in_group:
            result.append(self._separator)
            self.is_in_group = False
            self._bottom_only_pending = False
            self._just_closed = True
        elif has_bottom and not self.is_in_group:
            self.is_in_group = True
            self._bottom_only_pending = True

        return result

    def finalize(self) -> str | None:
        """If still in a group at end of document, emit closing separator."""
        if self.is_in_group:
            self.is_in_group = False
            self._bottom_only_pending = False
            return self._separator or None
        return None
