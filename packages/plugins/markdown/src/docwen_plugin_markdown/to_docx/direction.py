"""Resolve authored paragraph direction without treating alignment as bidi."""

from typing import Any

from docx.oxml.ns import qn


def style_is_rtl(style: Any) -> bool:
    seen: set[str] = set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        flag = style._element.find(f"{qn('w:pPr')}/{qn('w:bidi')}")
        if flag is not None:
            return flag.get(qn("w:val"), "1") not in {"0", "false", "off"}
        style = style.base_style
    return False


def paragraph_is_rtl(paragraph: Any) -> bool:
    flag = paragraph._p.find(f"{qn('w:pPr')}/{qn('w:bidi')}")
    if flag is not None:
        return flag.get(qn("w:val"), "1") not in {"0", "false", "off"}
    return style_is_rtl(paragraph.style)
