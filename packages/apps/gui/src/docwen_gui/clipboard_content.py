"""Project rich clipboard HTML into safe Markdown-oriented plain text.

The clipboard remains a text ingress. HTML is consulted only to preserve
table structure that can be represented without guessing. Unsupported table
shapes fall back to tab/newline-delimited cell text so user content is not
silently discarded.
"""

from __future__ import annotations

import html
import re
from dataclasses import dataclass
from html.parser import HTMLParser

_BLOCK_TAGS = frozenset(
    {
        "address",
        "article",
        "aside",
        "blockquote",
        "div",
        "footer",
        "h1",
        "h2",
        "h3",
        "h4",
        "h5",
        "h6",
        "header",
        "li",
        "main",
        "nav",
        "p",
        "section",
    }
)
_SKIP_TAGS = frozenset({"head", "script", "style"})


@dataclass(frozen=True, slots=True)
class ClipboardMarkdownProjection:
    text: str
    table_count: int = 0
    fallback_table_count: int = 0
    image_count: int = 0


@dataclass(frozen=True, slots=True)
class _TableCell:
    text: str
    header: bool
    rowspan: int = 1
    colspan: int = 1


def _positive_span(value: str | None) -> int:
    try:
        return max(1, min(int(value or "1"), 256))
    except (TypeError, ValueError):
        return 1


def _normalize_text(value: str) -> str:
    value = value.replace("\r\n", "\n").replace("\r", "\n").replace("\u00a0", " ")
    value = re.sub(r"[ \t\f\v]+", " ", value)
    value = re.sub(r" *\n *", "\n", value)
    value = re.sub(r"\n{3,}", "\n\n", value)
    return value.strip()


def _normalize_cell(value: str) -> str:
    value = value.replace("\r\n", "\n").replace("\r", "\n").replace("\u00a0", " ")
    value = re.sub(r"[ \t\f\v]+", " ", value)
    value = re.sub(r" *\n *", "\n", value)
    return value.strip()


def _markdown_cell(value: str) -> str:
    escaped = html.escape(value, quote=False).replace("\\", "\\\\").replace("|", "\\|")
    return escaped.replace("\n", "<br>")


class _ClipboardHTMLParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.blocks: list[str] = []
        self.text_parts: list[str] = []
        self.table_count = 0
        self.fallback_table_count = 0
        self.image_count = 0
        self._skip_depth = 0
        self._table_depth = 0
        self._table_nested = False
        self._rows: list[list[_TableCell]] = []
        self._row: list[_TableCell] | None = None
        self._cell_parts: list[str] | None = None
        self._cell_header = False
        self._cell_rowspan = 1
        self._cell_colspan = 1

    def _text_boundary(self) -> None:
        if not self.text_parts or not self.text_parts[-1].endswith("\n"):
            self.text_parts.append("\n")

    def _flush_text(self) -> None:
        value = _normalize_text("".join(self.text_parts))
        self.text_parts.clear()
        if value:
            self.blocks.append(value)

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.lower()
        if self._skip_depth:
            if tag in _SKIP_TAGS:
                self._skip_depth += 1
            return
        if tag in _SKIP_TAGS:
            if self.text_parts and self.text_parts[-1].endswith("\n"):
                self._flush_text()
            self._skip_depth = 1
            return

        if tag == "img":
            # Keep an existing paragraph boundary while dropping the image.
            # Inline images do not invent a paragraph break.
            if self.text_parts and self.text_parts[-1].endswith("\n"):
                self._flush_text()
            self.image_count += 1
            return

        if tag == "table":
            if self._table_depth == 0:
                self._flush_text()
                self._rows = []
                self._row = None
                self._cell_parts = None
                self._table_nested = False
            else:
                self._table_nested = True
            self._table_depth += 1
            return

        if self._table_depth:
            if self._table_depth == 1 and tag == "tr":
                self._row = []
            elif self._table_depth == 1 and tag in {"td", "th"} and self._row is not None:
                values = {name.lower(): value for name, value in attrs}
                self._cell_parts = []
                self._cell_header = tag == "th"
                self._cell_rowspan = _positive_span(values.get("rowspan"))
                self._cell_colspan = _positive_span(values.get("colspan"))
            elif tag == "br" and self._cell_parts is not None:
                self._cell_parts.append("\n")
            return

        if tag == "br":
            self.text_parts.append("\n")
        elif tag in _BLOCK_TAGS:
            self._text_boundary()

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if self._skip_depth:
            if tag in _SKIP_TAGS:
                self._skip_depth -= 1
            return

        if tag == "table" and self._table_depth:
            self._table_depth -= 1
            if self._table_depth == 0:
                self._finish_table()
            return

        if self._table_depth:
            if (
                self._table_depth == 1
                and tag in {"td", "th"}
                and self._cell_parts is not None
                and self._row is not None
            ):
                self._row.append(
                    _TableCell(
                        _normalize_cell("".join(self._cell_parts)),
                        self._cell_header,
                        self._cell_rowspan,
                        self._cell_colspan,
                    )
                )
                self._cell_parts = None
            elif self._table_depth == 1 and tag == "tr" and self._row is not None:
                if self._row:
                    self._rows.append(self._row)
                self._row = None
            return

        if tag in _BLOCK_TAGS:
            self._text_boundary()

    def handle_data(self, data: str) -> None:
        if self._skip_depth:
            return
        if self._table_depth:
            if self._cell_parts is not None:
                self._cell_parts.append(data)
            return
        self.text_parts.append(data)

    @staticmethod
    def _fallback_table_text(rows: list[list[_TableCell]]) -> str:
        lines: list[str] = []
        for row in rows:
            values: list[str] = []
            for cell in row:
                values.append(cell.text)
                values.extend("" for _ in range(cell.colspan - 1))
            lines.append("\t".join(values))
        return "\n".join(lines)

    def _finish_table(self) -> None:
        rows = [row for row in self._rows if row]
        self._rows = []
        if not rows:
            return

        width = len(rows[0])
        simple = (
            not self._table_nested
            and width > 0
            and len(rows) * width <= 65536
            and all(len(row) == width for row in rows)
            and all(cell.rowspan == 1 and cell.colspan == 1 for row in rows for cell in row)
        )
        explicit_header = simple and all(cell.header for cell in rows[0])
        if simple and explicit_header:
            rendered = [
                "| " + " | ".join(_markdown_cell(cell.text) for cell in rows[0]) + " |",
                "| " + " | ".join("---" for _ in range(width)) + " |",
            ]
            rendered.extend("| " + " | ".join(_markdown_cell(cell.text) for cell in row) + " |" for row in rows[1:])
            self.blocks.append("\n".join(rendered))
            self.table_count += 1
            return

        fallback = self._fallback_table_text(rows)
        if fallback:
            self.blocks.append(fallback)
        self.fallback_table_count += 1

    def projection(self) -> ClipboardMarkdownProjection:
        if self._table_depth:
            self._table_depth = 0
            self._finish_table()
        self._flush_text()
        value = "\n\n".join(block for block in self.blocks if block.strip()).strip("\n")
        if value:
            value += "\n"
        return ClipboardMarkdownProjection(value, self.table_count, self.fallback_table_count, self.image_count)


def project_clipboard_html(html_text: str) -> ClipboardMarkdownProjection:
    """Preserve reliable table structure while otherwise keeping HTML as plain text."""

    if not isinstance(html_text, str):
        raise TypeError("clipboard HTML must be text")
    parser = _ClipboardHTMLParser()
    parser.feed(html_text)
    parser.close()
    return parser.projection()


__all__ = ["ClipboardMarkdownProjection", "project_clipboard_html"]
