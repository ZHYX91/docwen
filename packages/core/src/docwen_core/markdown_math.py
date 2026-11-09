"""Shared dollar-math boundaries and Markdown serialization.

Block owners (the Markdown parser and the source-protection scanner) decide
which container owns a line. This module alone decides where its math starts
and ends; formula bodies are never parsed as Markdown.
"""

from __future__ import annotations

import re
from typing import TypedDict


class FormulaSpan(TypedDict):
    """A formula and its delimiter-inclusive source range."""

    latex: str
    is_inline: bool
    start: int
    end: int


DISPLAY_MATH_START = re.compile(r"^ {0,3}\$\$")


def _escaped(text: str, position: int) -> bool:
    preceding = position - 1
    while preceding >= 0 and text[preceding] == "\\":
        preceding -= 1
    return (position - preceding - 1) % 2 == 1


def display_math_closer(text: str, start: int = 0, *, standalone: bool = False) -> int | None:
    """Find the first unescaped ``$$`` closer, optionally at a line end."""
    position = text.find("$$", start)
    while position >= 0:
        if not _escaped(text, position):
            if not standalone:
                return position
            end_of_line = text.find("\n", position + 2)
            suffix = text[position + 2 : None if end_of_line < 0 else end_of_line]
            # A non-standalone pair belongs to an inline container. Do not
            # swallow it and a later block while searching for another pair.
            return position if not suffix.strip(" \t\r") else None
        position = text.find("$$", position + 2)
    return None


def match_formula(text: str, start: int = 0, *, standalone: bool = False) -> FormulaSpan | None:
    """Match dollar math at an already-qualified source position.

    Single dollars cannot cross a physical line. Opening whitespace, closing
    whitespace, and a digit after the closing dollar are excluded so ordinary
    amounts such as ``$5 and $10`` remain prose. Double dollars may span lines
    and retain their authored LaTeX line breaks.
    """
    if start >= len(text) or text[start] != "$" or _escaped(text, start):
        return None
    display = text.startswith("$$", start)
    if standalone and not display:
        return None
    if display:
        closing = display_math_closer(text, start + 2, standalone=standalone)
        if closing is None:
            return None
        end = closing + 2
        latex = text[start + 2 : closing].strip()
    else:
        if start and text[start - 1] == "$":
            return None
        body_start = start + 1
        if body_start == len(text) or text[body_start].isspace():
            return None
        closing = body_start
        while closing < len(text):
            if text[closing] in "\r\n":
                return None
            if text[closing] == "$" and not _escaped(text, closing):
                break
            closing += 1
        if closing == len(text) or closing == body_start or text[closing - 1].isspace():
            return None
        if text.startswith("$$", closing) or (closing + 1 < len(text) and text[closing + 1].isdigit()):
            return None
        end = closing + 1
        latex = text[body_start:closing]
    return {"latex": latex, "is_inline": not display, "start": start, "end": end}


def serialize_formula(latex: str, *, block: bool, compact: bool = False) -> str:
    """Emit re-importable dollar math without changing inline/display type."""
    latex = latex.strip()
    if compact or not block:
        latex = latex.replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
    if block:
        return f"$$\n{latex}\n$$" if "\n" in latex or "\r" in latex else f"$${latex}$$"
    # TeX whitespace is immaterial here; Markdown inline math cannot cross a
    # physical line. Matrix row separators (two backslashes) remain intact.
    return f"${latex}$"
