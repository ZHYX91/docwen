"""Readable projection of native external cell links, without evaluating formulas."""

from __future__ import annotations

from typing import Any
from urllib.parse import quote, urlsplit

from docwen_core.markdown_utils import encode_markdown_table_cell


def native_cell_markdown(cell: Any) -> str:
    """Keep the display value and a safe external relationship destination."""
    display = encode_markdown_table_cell(str(cell.value))
    hyperlink = getattr(cell, "hyperlink", None)
    target = getattr(hyperlink, "target", None)
    if not isinstance(target, str) or not target or any(ord(char) < 32 for char in target):
        return display
    try:
        parsed = urlsplit(target)
    except ValueError:
        return display
    scheme = parsed.scheme.lower()
    if scheme in ("http", "https", "ftp", "obsidian"):
        if not parsed.netloc:
            return display
    elif scheme in ("file", "mailto"):
        if not parsed.path:
            return display
    else:
        # Relative links and workbook locations need an independent source map.
        # HYPERLINK formulas are never parsed, evaluated or guessed here.
        return display
    destination = quote(target, safe=":/#?=&%+@;,~-._!")
    return f"[{display}](<{destination}>)"
