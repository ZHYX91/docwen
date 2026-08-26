"""Page/section break detection and border group tracking for DOCX->MD."""

from __future__ import annotations

from typing import Any

from docwen_core.docx_parsing.xml_ns import NS_W

# ── Break detection ─────────────────────────────────────────────────────


def detect_page_break(para: Any) -> bool:
    """Return True if the paragraph contains a page break."""
    nsmap = {"w": NS_W}
    breaks = para._p.findall(".//w:br", nsmap)

    def _is_in_rejected_revision(element: Any) -> bool:
        parent = element.getparent() if hasattr(element, "getparent") else None
        while parent is not None and parent is not para._p:
            local_name = parent.tag.split("}")[-1] if "}" in (parent.tag or "") else (parent.tag or "")
            if local_name in {"del", "moveFrom"}:
                return True
            parent = parent.getparent() if hasattr(parent, "getparent") else None
        return False

    return any(br.get(f"{{{NS_W}}}type") == "page" and not _is_in_rejected_revision(br) for br in breaks)


def detect_page_break_in_run(run: Any) -> bool:
    """Return True if a run contains a page break."""
    br = run._r.find(f"{{{NS_W}}}br")
    if br is not None:
        return br.get(f"{{{NS_W}}}type") == "page"
    return False


def detect_section_break(para: Any) -> tuple[str, str] | None:
    """Return (type, None) or None if no section break."""
    elem = para._p
    sect_pr = elem.find(f".//{{{NS_W}}}sectPr")
    if sect_pr is None:
        return None
    # Check for type attribute on sectPr
    stype = sect_pr.get(f"{{{NS_W}}}type") or "nextPage"
    return (stype, None)  # pyright: ignore[reportReturnType]


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
    if info.get("bottom") in ("single", "double"):
        text = para.text.strip() if hasattr(para, "text") else ""
        return text == ""
    return False


def extract_paragraph_border_info(para: Any) -> dict:
    """Extract paragraph border info dict (top/bottom/left/right/between)."""
    info: dict = {}
    try:
        pPr = para._p.find(f"{{{NS_W}}}pPr")
        if pPr is None:
            return info
        pBdr = pPr.find(f"{{{NS_W}}}pBdr")
        if pBdr is None:
            return info
        for direction in ("top", "bottom", "left", "right", "between"):
            bdr = pBdr.find(f"{{{NS_W}}}{direction}")
            if bdr is not None:
                val = bdr.get(f"{{{NS_W}}}val")
                if val and val not in ("none", "nil"):
                    info[direction] = val
    except Exception:
        pass
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


# ── Page‑break text splitting ─────────────────────────────────────────


def split_paragraph_by_page_breaks(para: Any, separator: str = "---") -> list[str]:
    """Split a paragraph's runs into text segments at page‑break positions.

    Each page break produces a separator between segments.
    If no page breaks exist the full paragraph text is returned as
    a single element.

    Args:
        para: python-docx Paragraph object.
        separator: The separator string to insert at break points.

    Returns:
        List of text segments with separator at split points.
    """
    result: list[str] = []
    current_buf: list[str] = []

    for run in para.runs:
        if detect_page_break_in_run(run):
            # Close current text segment
            text = "".join(current_buf).strip()
            if text:
                result.append(text)
            current_buf = []
            # Emit separator (collapse consecutive)
            if separator and (not result or result[-1] != separator):
                result.append(separator)
        else:
            if run.text:
                current_buf.append(run.text)

    # Final text segment
    text = "".join(current_buf).strip()
    if text:
        result.append(text)

    if not result:
        return [para.text]
    return result
