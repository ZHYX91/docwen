"""Built-in Gongwen field values and Word template layout.

YAML is source data.  Normalization returns a separate display projection;
attachment labels, numbering and indentation exist only in the Word output.
"""

from __future__ import annotations

import datetime as dt
import math
import re
from collections.abc import Mapping, Sequence
from copy import deepcopy
from typing import Any, Literal, cast

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.text.paragraph import CT_P
from docx.oxml.xmlchemy import BaseOxmlElement
from docx.shared import Emu, Pt, Twips
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from lxml import etree

from docwen_plugin_markdown.template_ownership import iter_paragraph_placeholder_spans
from docwen_plugin_markdown.template_parts import template_story_parts
from docwen_plugin_markdown.template_utils import effective_indent_attributes, inherited_indent_attributes
from docwen_plugin_markdown.yaml_processor import is_empty_field_value

_PLACEHOLDER = re.compile(r"\{\{\s*([^{}\r\n]+?)\s*\}\}")
_ATTACHMENT = re.compile(r"\{\{\s*附件说明\s*\}\}")
_ATTACHMENT_PLACEMENT_MESSAGE = (
    "附件说明占位符必须独占正文或表格单元格中的普通段落；"
    "不支持跨换行、制表符或其他内容边界，或放在页眉、页脚、文本框等位置。"
)
_ATTACHMENT_NUMBER = re.compile(
    r"^(?:[一二三四五六七八九十百]+[、.]|"
    r"[（(](?:[一二三四五六七八九十百]+|\d+)[）)]|"
    r"\d+[.．、]|"
    r"[㈠㈡㈢㈣㈤㈥㈦㈧㈨㈩①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"
    r"㉑㉒㉓㉔㉕㉖㉗㉘㉙㉚㉛㉜㉝㉞㉟㊱㊲㊳㊴㊵㊶㊷㊸㊹㊺㊻㊼㊽㊾㊿"
    r"⓪⓵⓶⓷⓸⓹⓺⓻⓼⓽⓾⓿❶❷❸❹❺❻❼❽❾❿⓫⓬⓭⓮⓯⓰⓱⓲⓳⓴])\s*"
)
_PARAGRAPH_GROUPS = (
    ("密级和保密期限",),
    ("紧急程度",),
    ("发文字号",),
    ("公开方式",),
    ("主送机关",),
    ("签发人",),
    ("附注",),
    ("抄送机关",),
    ("附件说明",),
    ("份号", "发文字号"),
)
_ROW_GROUPS = (("抄送机关",), ("印发机关", "印发日期"))
_LABELS = {"抄送机关": ("抄送",), "印发机关": ("版记",), "附件说明": ("附件",)}
_DECORATION = re.compile(r"[\s:：;；,，、.。()（）\[\]【】/／—\-]*\Z")
# Deleting an empty field is a layout operation. Only ordinary text and known
# formatting may disappear with its container; unknown extensions, revision
# records and semantic anchors must remain attached even when their text is empty.
_LAYOUT_CHILDREN = {
    "p": "pPr r",
    "r": "rPr t tab",
    "tr": "trPr tc",
    "tc": "tcPr p",
    "tbl": "tblPr tblGrid tr",
    "tblGrid": "gridCol",
    "pPr": (
        "pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs "
        "suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd "
        "snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment "
        "textboxTightWrap outlineLvl divId cnfStyle rPr"
    ),
    "rPr": (
        "rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid "
        "vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs "
        "em lang eastAsianLayout specVanish oMath"
    ),
    "trPr": "cnfStyle divId gridBefore gridAfter wBefore wAfter cantSplit trHeight tblHeader tblCellSpacing jc hidden",
    "tcPr": "cnfStyle tcW gridSpan tcBorders shd noWrap tcMar textDirection tcFitText vAlign hideMark",
    "tblPr": (
        "tblStyle tblpPr tblOverlap bidiVisual tblStyleRowBandSize tblStyleColBandSize tblW jc tblCellSpacing "
        "tblInd tblBorders shd tblLayout tblCellMar tblLook"
    ),
    "numPr": "ilvl numId",
    "tabs": "tab",
    "pBdr": "top left bottom right between bar",
    "tcBorders": "top left bottom right start end insideH insideV tl2br tr2bl",
    "tblBorders": "top left bottom right start end insideH insideV",
    "tcMar": "top left bottom right start end",
    "tblCellMar": "top left bottom right start end",
}
_EMPTY_LAYOUT_CHILDREN = {
    qn(f"w:{name}"): frozenset(qn(f"w:{child}") for child in children.split())
    for name, children in _LAYOUT_CHILDREN.items()
}
for _children in tuple(_EMPTY_LAYOUT_CHILDREN.values()):
    for _child in _children:
        _EMPTY_LAYOUT_CHILDREN.setdefault(_child, frozenset())
del _children, _child

# Core Word attributes describe the recognized layout nodes above. Only the
# paragraph identity pair is disposable from the Word 2010 extension namespace;
# other extension attributes can carry semantics even without child elements.
_WORD_ATTRIBUTE_NAMESPACE = qn("w:val").removesuffix("val")
_PARAGRAPH_IDENTITY_ATTRIBUTES = frozenset({qn("w14:paraId"), qn("w14:textId")})


class GongwenFieldPlacementError(ValueError):
    """A Gongwen field cannot be safely rendered or reviewed at its location."""

    diagnostic_code = "MD2DOCX-GONGWEN-FIELD-PLACEMENT"


def prepare_gongwen_fields(source: Mapping[str, Any]) -> dict[str, Any]:
    """Build display values without changing the caller's YAML or nested data."""
    values = dict(source)
    if "附件说明" in source:
        values["附件说明"] = _attachment_names(source["附件说明"])
    for key in ("签发人", "主送机关", "抄送机关"):
        if key in source:
            value = source[key]
            values[key] = value if isinstance(value, str) else "、".join(_field_list_items(key, value))
    if "附注" in source:
        values["附注"] = _notes(source["附注"])
    for key, suffix in (("成文日期", ""), ("印发日期", "印发")):
        if key in source:
            values[key] = _date(source[key], suffix)
    return values


def _field_list_items(key: str, value: Any) -> tuple[str, ...]:
    """Accept a flat list of authored strings; never infer names or categories."""
    if value is None:
        return ()
    if not isinstance(value, (list, tuple)):
        raise ValueError(f"“{key}”请填写一段文字，或每项都是文字的列表。")
    items = []
    for item in value:
        if item is None:
            continue
        if not isinstance(item, str):
            raise ValueError(f"“{key}”的每一项都应是文字，不支持嵌套列表、数字或对象。")
        if item.strip():
            items.append(item.strip())
    return tuple(items)


def gongwen_review_notes(source: Mapping[str, Any]) -> dict[str, str]:
    """Explain only multi-item list rendering, not already-authored punctuation."""
    notes = {}
    for key in ("签发人", "主送机关", "抄送机关"):
        value = source.get(key)
        if isinstance(value, str) or len(_field_list_items(key, value)) < 2:
            continue
        if key == "签发人":
            notes[key] = (
                "这里有多位签发人，已暂用顿号连接。请在 Word 中自行调整姓名的换行和对齐，"
                "并核对姓名顺序；本程序没有自动排版。"
            )
        else:
            notes[key] = f"{key}有多项，已暂用顿号连接。请核对哪些地方应使用顿号、哪些应改成逗号，并检查机关排列顺序。"
    return notes


def validate_review_placeholders(doc: Any, placeholders: Mapping[str, Sequence[Any]], notes: Mapping[str, str]) -> None:
    """Do not silently drop required review comments in unsupported Word parts."""
    for key in notes:
        for paragraph in placeholders.get(key, ()):
            ancestor = paragraph._p.getparent()
            while ancestor is not None and ancestor is not doc.element.body:
                if ancestor.tag not in {qn("w:tc"), qn("w:tr"), qn("w:tbl")}:
                    break
                ancestor = ancestor.getparent()
            if ancestor is not doc.element.body:
                raise GongwenFieldPlacementError(
                    f"“{key}”有多个条目，需要添加核对批注。请将占位符放在正文或表格中，"
                    "或改用已核对标点和排版的单段文字。"
                )


def validate_attachment_placeholders(
    doc: Any,
    placeholders: Mapping[str, Sequence[Any]],
    *,
    rendered_elements: Sequence[etree._Element] = (),
) -> None:
    """Accept standalone, ordinary paragraphs in the main body or table cells."""
    _validate_reserved_attachment_markers(doc, rendered_elements)
    for paragraph in placeholders.get("附件说明", ()):
        element = paragraph._p
        ancestor = element.getparent()
        allowed = ancestor is doc.element.body or (ancestor is not None and ancestor.tag == qn("w:tc"))
        while ancestor is not None and ancestor is not doc.element.body:
            if ancestor.tag not in {qn("w:tc"), qn("w:tr"), qn("w:tbl")}:
                allowed = False
                break
            ancestor = ancestor.getparent()
        # Restrict the anchor, not inserted YAML values or rendered Markdown.
        if (
            not allowed
            or ancestor is not doc.element.body
            or _ATTACHMENT.fullmatch(paragraph.text.strip()) is None
            or not _plain_layout_container(element, allow_section=True)
        ):
            raise GongwenFieldPlacementError(_ATTACHMENT_PLACEMENT_MESSAGE)


def _validate_reserved_attachment_markers(doc: Any, rendered_elements: Sequence[etree._Element]) -> None:
    """Diagnose broken reserved markers without turning them into fill spans.

    This deliberately wider view ignores visible control boundaries solely for
    diagnosis. Only the shared span reader defines replaceable template text.
    Rendered Markdown is literal input and must never become a template marker.
    """
    roots = [part.element for part in template_story_parts(doc)]
    ignored = frozenset(rendered_elements)
    for root in roots:
        for paragraph in root.iter(qn("w:p")):
            if paragraph in ignored or any(ancestor in ignored for ancestor in paragraph.iterancestors()):
                continue
            valid = {span.locations for span in iter_paragraph_placeholder_spans(paragraph) if span.key == "附件说明"}
            text, locations = _reserved_marker_text(paragraph)
            for match in _ATTACHMENT.finditer(text):
                if tuple(locations[match.start() : match.end()]) not in valid:
                    raise GongwenFieldPlacementError(_ATTACHMENT_PLACEMENT_MESSAGE)


def _reserved_marker_text(element: etree._Element) -> tuple[str, list[tuple[etree._Element, int]]]:
    """Read the current paragraph's raw text for conservative diagnostics."""
    inactive = {qn("w:del"), qn("w:moveFrom")}
    if any(ancestor.tag in inactive for ancestor in element.iterancestors()):
        return "", []
    characters: list[str] = []
    locations: list[tuple[etree._Element, int]] = []

    def visit(node: etree._Element) -> None:
        if node.tag in inactive or (node is not element and node.tag == qn("w:p")):
            return
        if node.tag == qn("w:t"):
            for offset, character in enumerate(node.text or ""):
                if character not in "\t\r\n":
                    characters.append(character)
                    locations.append((node, offset))
        else:
            for child in node:
                visit(child)

    visit(element)
    return "".join(characters), locations


def remove_empty_gongwen_fields(doc: Any, values: Mapping[str, Any], placeholders: Mapping[str, Sequence[Any]]) -> None:
    """Remove only empty field-only containers from the original template."""
    for group in _ROW_GROUPS:
        for row in _matching_containers(placeholders, group, "tr"):
            if not _removable_container(row, group, values):
                continue
            table = row.getparent()
            if table is None:
                continue
            if len(table.findall(qn("w:tr"))) == 1 and not _plain_layout_container(table):
                continue
            table.remove(row)
            if not table.findall(qn("w:tr")):
                parent = table.getparent()
                if parent is not None:
                    parent.remove(table)
                    _ensure_cell_paragraph(parent)
    for group in _PARAGRAPH_GROUPS:
        for paragraph in _matching_containers(placeholders, group, "p"):
            if _removable_container(paragraph, group, values):
                _remove_paragraph(paragraph)


def render_attachment_placeholders(
    names: Sequence[str],
    placeholders: Mapping[str, Sequence[Any]],
    *,
    alignment: Literal["spaces", "tab"] = "spaces",
) -> None:
    """Expand each original attachment anchor in its own body/cell container."""
    for target in placeholders.get("附件说明", ()):
        parent = target._p.getparent()
        if parent is None:
            continue
        if not names:
            # A section-bearing empty paragraph must keep its section break.
            for child in list(target._p):
                if child.tag != qn("w:pPr"):
                    target._p.remove(child)
            _remove_paragraph(target._p)
            continue

        source_run = next((run for run in target.runs if "{{" in run.text or "{" in run.text), None)
        if source_run is None:
            source_run = next(iter(target.runs), None)
        char_width = _font_size(target, source_run)
        inherited = inherited_indent_attributes(target)
        base_indent = _attachment_origin(target, char_width)
        # Keep the period and following ordinary space adjacent to the name.
        # Reserve the label's six half-character cells, reducing the leading
        # padding by one cell whenever the number gains a digit.
        source_hanging = _indent_length(effective_indent_attributes(target), char_width, "hanging")
        number_cells = 0 if len(names) == 1 else 2
        padding_cells = max(6, len(str(len(names))) - 1, math.ceil(source_hanging * 2 / char_width) - number_cells)
        hanging = int((padding_cells + number_cells) * char_width / 2)
        text_indent = base_indent + hanging
        insertion_index = parent.index(target._p)

        for index, name in enumerate(names, 1):
            element = cast(CT_P, OxmlElement("w:p"))
            if target._p.pPr is not None:
                element.append(deepcopy(target._p.pPr))
                # A section boundary belongs after the complete expansion.
                properties = element.get_or_add_pPr()
                section = properties.find(qn("w:sectPr"))
                if section is not None and index < len(names):
                    properties.remove(section)
            paragraph = Paragraph(element, target._parent)
            run = paragraph.add_run()
            if source_run is not None and source_run._r.rPr is not None:
                run._r.insert(0, deepcopy(source_run._r.rPr))
            if len(names) == 1:
                run.text = f"附件：{_attachment_padding(padding_cells - 6)}"
                if alignment == "spaces":
                    run.text += name
                else:
                    run.add_tab()
                    name_run = paragraph.add_run(name)
                    if source_run is not None and source_run._r.rPr is not None:
                        name_run._r.insert(0, deepcopy(source_run._r.rPr))
            else:
                prefix = (
                    "附件：" + _attachment_padding(padding_cells - 6)
                    if index == 1
                    else _attachment_padding(padding_cells - len(str(index)) + 1)
                )
                run.text = f"{prefix}{index}." + (" " if alignment == "spaces" else "\t")
                # A template can kern pairs such as "11" more tightly. Keep
                # literal number cells uniform without changing the name's font.
                properties = run._r.get_or_add_rPr()
                kern = properties.find(qn("w:kern"))
                if kern is None:
                    new_kern = cast(BaseOxmlElement, OxmlElement("w:kern"))
                    properties.insert_element_before(
                        new_kern,
                        "w:position",
                        "w:sz",
                        "w:szCs",
                        "w:highlight",
                        "w:u",
                        "w:effect",
                        "w:bdr",
                        "w:shd",
                        "w:fitText",
                        "w:vertAlign",
                        "w:rtl",
                        "w:cs",
                        "w:em",
                        "w:lang",
                        "w:eastAsianLayout",
                        "w:specVanish",
                        "w:oMath",
                        "w:rPrChange",
                    )
                    kern = new_kern
                kern.set(qn("w:val"), str(int(Emu(char_width).pt * 2) + 1))
                name_run = paragraph.add_run(name)
                if source_run is not None and source_run._r.rPr is not None:
                    name_run._r.insert(0, deepcopy(source_run._r.rPr))
            paragraph.paragraph_format.left_indent = text_indent
            paragraph.paragraph_format.first_line_indent = -hanging
            # Character-based indents would override the physical layout above.
            indent = paragraph._p.get_or_add_pPr().get_or_add_ind()
            for attribute in ("leftChars", "startChars", "firstLineChars", "hangingChars", "start"):
                key = qn(f"w:{attribute}")
                if key in indent.attrib:
                    del indent.attrib[key]
            # Removing a direct value exposes style/default character units.
            # Word requires a local zero to cancel those earlier declarations;
            # keep the template's shared styles and unrelated axes untouched.
            for attribute in ("leftChars", "startChars"):
                key = qn(f"w:{attribute}")
                if int(inherited.get(key, "0")):
                    indent.set(key, "0")
            if any(int(inherited.get(qn(f"w:{name}"), "0")) for name in ("firstLineChars", "hangingChars")):
                indent.set(qn("w:hangingChars"), "0")
            if any(qn(f"w:{name}") in inherited for name in ("start", "startChars")):
                indent.set(qn("w:start"), str(Emu(text_indent).twips))
            if alignment == "tab":
                _attachment_name_tab(paragraph, text_indent)
            parent.insert(insertion_index + index - 1, element)
        parent.remove(target._p)


def _attachment_name_tab(paragraph: Paragraph, position: int) -> None:
    """Use one name stop, cancelling only earlier authored stops locally.

    Stops are inherited independently of indentation. An earlier template stop
    would capture our single tab before it reached the continuation origin.
    Shared styles and later unrelated stops must remain intact.
    """
    effective: dict[int, str] = {}
    properties: list[Any] = [paragraph._p.pPr]
    style = paragraph.style
    seen: set[Any] = set()
    while style is not None and style.element not in seen:
        seen.add(style.element)
        properties.append(style.element.find(qn("w:pPr")))
        style = style.base_style
    properties.append(
        cast(Any, paragraph.part).styles.element.find(f"{qn('w:docDefaults')}/{qn('w:pPrDefault')}/{qn('w:pPr')}")
    )
    destination = Emu(position).twips
    for source in properties:
        if source is None:
            continue
        tabs = source.find(qn("w:tabs"))
        if tabs is not None:
            for stop in tabs:
                effective.setdefault(int(stop.get(qn("w:pos"), "0")), stop.get(qn("w:val"), "left"))
    original_destination = destination
    while effective.get(destination) == "bar":
        destination += 1
    if destination != original_destination:
        # A visible bar and a positioning stop cannot share one declaration.
        # Move only the generated name origin by the smallest physical unit.
        difference = Twips(destination - original_destination)
        first_line = paragraph.paragraph_format.first_line_indent
        assert first_line is not None
        paragraph.paragraph_format.left_indent = Twips(destination)
        paragraph.paragraph_format.first_line_indent = Emu(first_line - difference)
        indent = paragraph._p.get_or_add_pPr().get_or_add_ind()
        if qn("w:start") in indent.attrib:
            indent.set(qn("w:start"), str(destination))
        position = Twips(destination)
    tabs = cast(Any, paragraph._p.get_or_add_pPr()).get_or_add_tabs()
    for stop in list(tabs):
        stop_position = int(stop.get(qn("w:pos"), "0"))
        kind = stop.get(qn("w:val"))
        if stop_position <= destination and kind != "bar" and (kind != "clear" or stop_position == destination):
            tabs.remove(stop)
    for earlier in sorted(
        value for value, kind in effective.items() if value < destination and kind not in {"bar", "clear"}
    ):
        stop = OxmlElement("w:tab")
        stop.set(qn("w:val"), "clear")
        stop.set(qn("w:pos"), str(earlier))
        tabs.append(stop)
    paragraph.paragraph_format.tab_stops.add_tab_stop(Emu(position))
    ordered = sorted(tabs, key=lambda stop: int(stop.get(qn("w:pos"), "0")))
    for stop in ordered:
        tabs.remove(stop)
    for stop in ordered:
        tabs.append(stop)


def _attachment_names(value: Any) -> tuple[str, ...]:
    names = []
    for item in _display_items(value):
        text = re.sub(r"^附件\s*[:：]\s*", "", item)
        cleaned = _ATTACHMENT_NUMBER.sub("", text, count=1).strip()
        names.append(cleaned or text)
    return tuple(names)


def _display_items(value: Any):
    if is_empty_field_value(value):
        return
    if isinstance(value, (list, tuple)):
        for item in value:
            yield from _display_items(item)
    else:
        yield str(value).strip()


def _notes(value: Any) -> str:
    if is_empty_field_value(value):
        return ""
    if isinstance(value, (Mapping, list, tuple, set)):
        raise ValueError("“附注”请填写一段文字，不要填写列表或对象。")
    text = str(value).strip()
    if len(text) < 2 or (text[0], text[-1]) not in {("（", "）"), ("(", ")")}:
        return text
    depth = 0
    for index, character in enumerate(text):
        if character == text[0]:
            depth += 1
        elif character == text[-1]:
            depth -= 1
        if depth == 0 and index < len(text) - 1:
            return text
    return text[1:-1] if depth == 0 else text


def _date(value: Any, suffix: str) -> str:
    if is_empty_field_value(value):
        return ""
    if isinstance(value, (Mapping, list, tuple, set)):
        raise ValueError("成文日期和印发日期应填写单个日期或文字，不要填写列表或对象。")
    if isinstance(value, dt.date):
        parsed = value
    else:
        parsed = None
        for format_string in ("%Y-%m-%d", "%Y/%m/%d", "%Y年%m月%d日", "%Y.%m.%d", "%Y年%m月%d号"):
            try:
                parsed = dt.datetime.strptime(str(value).strip(), format_string)
                break
            except ValueError:
                continue
    return f"{parsed.year}年{parsed.month}月{parsed.day}日{suffix}" if parsed is not None else str(value)


def _matching_containers(placeholders: Mapping[str, Sequence[Any]], group: tuple[str, ...], tag: str) -> set[Any]:
    common = None
    for key in group:
        containers = set()
        for paragraph in placeholders.get(key, ()):
            element = paragraph._p
            while element is not None and element.tag != qn(f"w:{tag}"):
                element = element.getparent()
            if element is not None:
                containers.add(element)
        common = containers if common is None else common & containers
    return common or set()


def _removable_container(container: Any, group: tuple[str, ...], values: Mapping[str, Any]) -> bool:
    if any(not is_empty_field_value(values.get(key)) for key in group):
        return False
    if not _plain_layout_container(container):
        return False
    # Never delete a containing row because of fields in its nested table.
    if container.tag == qn("w:tr") and any(node.tag == qn("w:tbl") for node in container.iter()):
        return False
    paragraphs = [container] if container.tag == qn("w:p") else container.iter(qn("w:p"))
    labels = set(group)
    labels.update(label for key in group for label in _LABELS.get(key, ()))
    for paragraph in paragraphs:
        text = "".join(node.text or "" for node in paragraph.iter(qn("w:t")))
        fields = {match.group(1).strip() for match in _PLACEHOLDER.finditer(text)}
        if any(not is_empty_field_value(values.get(key)) for key in fields):
            return False
        remaining = _PLACEHOLDER.sub("", text)
        for label in sorted(labels, key=len, reverse=True):
            remaining = remaining.replace(label, "")
        if _DECORATION.fullmatch(remaining) is None:
            return False
    return True


def _plain_layout_container(container: Any, *, allow_section: bool = False) -> bool:
    """Recognize an ordinary text/layout tree, never an unknown semantic tree."""
    for ancestor in container.iterancestors():
        if ancestor.tag in {qn("w:body"), qn("w:hdr"), qn("w:ftr")}:
            break
        if ancestor.tag not in {qn("w:tbl"), qn("w:tr"), qn("w:tc")} or not _plain_layout_attributes(ancestor):
            return False
    else:
        return False
    pending = [container]
    while pending:
        node = pending.pop()
        allowed = _EMPTY_LAYOUT_CHILDREN.get(node.tag)
        if allowed is None or not _plain_layout_attributes(node):
            return False
        for child in node:
            if allow_section and node.tag == qn("w:pPr") and child.tag == qn("w:sectPr"):
                # Attachment expansion preserves this original section subtree.
                continue
            if child.tag not in allowed:
                return False
            pending.append(child)
    return True


def _plain_layout_attributes(node: Any) -> bool:
    """Permit core layout attributes, paragraph identities and text whitespace."""
    return all(
        attribute.startswith(_WORD_ATTRIBUTE_NAMESPACE)
        or (node.tag == qn("w:p") and attribute in _PARAGRAPH_IDENTITY_ATTRIBUTES)
        or (node.tag == qn("w:t") and attribute == qn("xml:space"))
        for attribute in node.attrib
    )


def _remove_paragraph(element: Any) -> None:
    parent = element.getparent()
    if parent is None or element.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None:
        return
    parent.remove(element)
    _ensure_cell_paragraph(parent)


def _ensure_cell_paragraph(element: Any) -> None:
    # Word requires a cell to finish with a paragraph, including after a table.
    if element.tag == qn("w:tc") and (not len(element) or element[-1].tag != qn("w:p")):
        element.append(OxmlElement("w:p"))


def _font_size(paragraph: Any, run: Run | None) -> int:
    if run is not None and run.font.size is not None:
        return int(run.font.size)
    styles = [run.style if run is not None else None, paragraph.style]
    for style in styles:
        while style is not None:
            if style.font.size is not None:
                return int(style.font.size)
            style = style.base_style
    return int(Pt(16))


def _attachment_origin(paragraph: Any, char_width: int) -> int:
    attributes = effective_indent_attributes(paragraph)
    axis = "start" if any(qn(f"w:{name}") in attributes for name in ("start", "startChars")) else "left"
    left = _indent_length(attributes, char_width, axis)
    if any(qn(f"w:{name}") in attributes for name in ("hanging", "firstLine")):
        first = "hanging" if qn("w:hanging") in attributes else "firstLine"
    elif int(attributes.get(qn("w:hangingChars"), "0")):
        first = "hanging"
    elif int(attributes.get(qn("w:firstLineChars"), "0")):
        first = "firstLine"
    else:
        first = "hanging" if qn("w:hanging") in attributes else "firstLine"
    offset = _indent_length(attributes, char_width, first)
    return left + (-offset if first == "hanging" else offset)


def _indent_length(attributes: Mapping[str, str], char_width: int, axis: str) -> int:
    # Word-authored paired values contain measured physical geometry. Replacing
    # that geometry with font-size arithmetic moves real template anchors when
    # character metadata and lengths disagree. Only character-only templates
    # need a font-based fallback. This policy is local to attachment layout.
    length = attributes.get(qn(f"w:{axis}"))
    if length is not None:
        return int(Twips(int(length)))
    characters = int(attributes.get(qn(f"w:{axis}Chars"), "0"))
    if characters:
        return round(characters * char_width / 100)
    return 0


def _attachment_padding(half_characters: int) -> str:
    # An EN SPACE is half an em, unlike a proportional-font U+0020. Only this
    # leading padding uses it; the separator after the period is always U+0020.
    full, half = divmod(half_characters, 2)
    return "\u3000" * full + "\u2002" * half
