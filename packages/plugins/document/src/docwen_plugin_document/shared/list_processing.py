"""Word list detection, counter management."""

from __future__ import annotations

from typing import Any

from docwen_core.docx_parsing.xml_ns import NS_W

# ── List counter manager ────────────────────────────────────────────────


class ListCounterManager:
    """Track list item counters per numId and level.

    Resets deeper levels when a parent level increments.
    """

    def __init__(self) -> None:
        self._counters: dict[str, dict[int, int]] = {}
        self._last_level: dict[str, int] = {}

    def next(self, num_id: str, level: int, *, start: int = 1) -> int:
        """Return the next counter value for (num_id, level).

        Automatically resets deeper levels when a shallower level increments.
        """
        if num_id not in self._counters:
            self._counters[num_id] = {}
            self._last_level[num_id] = -1

        prev_level = self._last_level[num_id]

        # Reset deeper levels when moving back to a shallower level
        if level <= prev_level:
            for lvl in list(self._counters[num_id].keys()):
                if lvl > level:
                    del self._counters[num_id][lvl]

        # Initialize this level if needed
        if level not in self._counters[num_id]:
            self._counters[num_id][level] = start - 1

        self._counters[num_id][level] += 1
        self._last_level[num_id] = level
        return self._counters[num_id][level]

    def peek(self, num_id: str, level: int, *, start: int = 1) -> int:
        """Return the next value without mutating counter state."""
        current = self._counters.get(num_id, {}).get(level)
        return start if current is None else current + 1

    def snapshot(self, num_id: str) -> dict[int, int]:
        """Return a detached level/value snapshot for placeholder rendering."""
        return dict(self._counters.get(num_id, {}))

    def reset(self, num_id: str | None = None) -> None:
        """Reset all counters, or counters for a specific num_id."""
        if num_id is None:
            self._counters.clear()
            self._last_level.clear()
        else:
            self._counters.pop(num_id, None)
            self._last_level.pop(num_id, None)


# ── List detection ──────────────────────────────────────────────────────


def detect_list_item(
    para: Any,
    numbering_index: Any = None,
) -> tuple[str | None, int | None, str | None]:
    """Detect if a paragraph is a Word list item.

    Args:
        para: python-docx Paragraph object.
        numbering_index: Optional NumberingIndex for resolving numPr.

    Returns:
        (num_id, ilvl, list_type) or (None, None, None) if not a list item.
    """
    w_ns = NS_W
    try:
        pPr = para._p.find(f"{{{w_ns}}}pPr")
        if pPr is None:
            return None, None, None
        numPr = pPr.find(f"{{{w_ns}}}numPr")
        if numPr is None:
            # Fallback: resolve via paragraph style → numbering definition
            if numbering_index is not None and para.style is not None:
                style_id = getattr(para.style, "style_id", None)
                if isinstance(style_id, str) and style_id:
                    level_info = numbering_index.lookup_by_style_id(style_id)
                    if level_info is not None:
                        list_type = "bullet" if level_info.num_fmt == "bullet" else "ordered"
                        return f"abs_{level_info.abstract_num_id}", level_info.ilvl, list_type
            return None, None, None

        num_id_elem = numPr.find(f"{{{w_ns}}}numId")
        ilvl_elem = numPr.find(f"{{{w_ns}}}ilvl")
        if num_id_elem is None:
            return None, None, None

        num_id = num_id_elem.get(f"{{{w_ns}}}val") or "0"
        # ``numId=0`` is the OOXML sentinel for explicitly disabling
        # numbering on a paragraph.  It is common on body paragraphs that
        # follow a list and must not fall through to the default bullet type.
        if num_id == "0":
            return None, None, None
        ilvl = int(ilvl_elem.get(f"{{{w_ns}}}val") or "0")

        # Determine list type (bullet vs numbered)
        list_type = "bullet"  # default
        if numbering_index is not None:
            level = numbering_index.lookup(num_id, ilvl)
            if level is not None:
                fmt = level.num_fmt.lower() if level.num_fmt else ""
                list_type = "bullet" if fmt in ("bullet",) else "ordered"

        return num_id, ilvl, list_type
    except Exception:
        return None, None, None


def format_list_marker(
    list_type: str,
    ordinal: int,
    marker_type: str = "-",
) -> str:
    """Format a list marker string for Markdown."""
    if list_type == "ordered":
        return f"{ordinal}."
    markers = {"dash": "-", "asterisk": "*", "plus": "+", "-": "-", "*": "*", "+": "+"}
    return markers.get(marker_type, "-")
