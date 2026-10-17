"""Safe, offline text projection for raw Markdown HTML blocks."""

from __future__ import annotations

import re
from html.parser import HTMLParser

_BLOCK_TAGS = frozenset(
    "article aside blockquote div dl dt dd h1 h2 h3 h4 h5 h6 "
    "header footer li main ol p pre section table tbody thead "
    "tfoot tr ul"
    .split()
)
_IGNORED_TAGS = frozenset({"script", "style", "iframe", "object", "svg", "template"})
_SIMPLE_INLINE_TAGS = frozenset("a b em i mark s small span strong sub sup u".split())
_WHITESPACE = re.compile(r"[ \t\r\n]+")


class _VisibleHtml(HTMLParser):
    """Extract display text without fetching resources or executing markup."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.lines: list[str] = [""]
        self.hidden_depth = 0
        self.pre_depth = 0
        self.lossy = False

    def _boundary(self) -> None:
        if self.lines[-1].strip():
            self.lines.append("")

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if self.hidden_depth:
            if tag in _IGNORED_TAGS:
                self.hidden_depth += 1
            return
        if tag in _IGNORED_TAGS:
            self.lossy = True
            self.hidden_depth = 1
        elif tag == "br":
            self._boundary()
        elif tag == "img":
            self.lossy = True
            alt = next((value for key, value in attrs if key == "alt"), None)
            if alt:
                self.handle_data(alt)
        elif tag == "pre":
            self.lossy = True
            self.pre_depth += 1
            self._boundary()
        elif tag in {"td", "th"}:
            self.lossy = True
            self._boundary()
        elif tag in _BLOCK_TAGS:
            if tag == "table":
                self.lossy = True
            self._boundary()
        elif tag not in _SIMPLE_INLINE_TAGS:
            self.lossy = True

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        self.handle_starttag(tag, attrs)
        self.handle_endtag(tag)

    def handle_endtag(self, tag: str) -> None:
        if self.hidden_depth:
            if tag in _IGNORED_TAGS:
                self.hidden_depth -= 1
            return
        if tag == "pre":
            self.pre_depth = max(0, self.pre_depth - 1)
            self._boundary()
        elif tag in _BLOCK_TAGS:
            self._boundary()

    def handle_data(self, data: str) -> None:
        if self.hidden_depth:
            return
        if self.pre_depth:
            for index, part in enumerate(data.split("\n")):
                if index:
                    self._boundary()
                self.lines[-1] += part
        else:
            normalized = _WHITESPACE.sub(" ", data)
            if self.lines[-1] == "":
                normalized = normalized.lstrip(" ")
            self.lines[-1] += normalized


def project_html_block_text(raw: str) -> tuple[list[str], bool]:
    """Return visible paragraphs and whether HTML was degraded or suppressed."""
    parser = _VisibleHtml()
    parser.feed(raw)
    parser.close()
    return [line.strip(" 	") for line in parser.lines if line.strip()], parser.lossy
