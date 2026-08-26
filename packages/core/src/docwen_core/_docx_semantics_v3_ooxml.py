"""Small OOXML mutation and recovery primitives for semantics v3."""

from __future__ import annotations

import re
from typing import Any

from docwen_core._docx_semantics_v3_model import DocxSemanticsV3Error

_REF_INSTRUCTION_RE = re.compile(
    r"\s*REF\s+(DW_T_[0-9a-f]{35})(?:\s+(\\n))?\s+\\h\s*",
    re.IGNORECASE,
)


def wrap_paragraph_content_with_bookmark(paragraph: Any, name: str, bookmark_id: str) -> None:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), bookmark_id)
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), bookmark_id)
    paragraph_element = paragraph._p
    insert_at = 1 if paragraph_element.pPr is not None else 0
    paragraph_element.insert(insert_at, start)
    paragraph_element.append(end)


def append_complex_field(paragraph: Any, *, instruction: str, cached_result: str) -> None:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    begin_run = OxmlElement("w:r")
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    begin.set(qn("w:dirty"), "true")
    begin_run.append(begin)
    paragraph._p.append(begin_run)
    instruction_run = OxmlElement("w:r")
    instruction_text = OxmlElement("w:instrText")
    instruction_text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instruction_text.text = instruction
    instruction_run.append(instruction_text)
    paragraph._p.append(instruction_run)
    separate_run = OxmlElement("w:r")
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    separate_run.append(separate)
    paragraph._p.append(separate_run)
    result_run = OxmlElement("w:r")
    result = OxmlElement("w:t")
    result.text = cached_result
    result_run.append(result)
    paragraph._p.append(result_run)
    end_run = OxmlElement("w:r")
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    end_run.append(end)
    paragraph._p.append(end_run)


def inline_sdt(tag: str, visible: str) -> Any:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    sdt = OxmlElement("w:sdt")
    properties = OxmlElement("w:sdtPr")
    tag_element = OxmlElement("w:tag")
    tag_element.set(qn("w:val"), tag)
    properties.append(tag_element)
    content = OxmlElement("w:sdtContent")
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    if visible[:1].isspace() or visible[-1:].isspace():
        text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text.text = visible
    run.append(text)
    content.append(run)
    sdt.extend((properties, content))
    return sdt


def inline_reference_sdt(
    tag: str,
    *,
    bookmark_name: str,
    cached_number: str,
    heading_number_only: bool,
    alias: str | None,
) -> Any:
    """Create one occurrence SDT containing an exact REF and optional Alias."""

    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    sdt = OxmlElement("w:sdt")
    properties = OxmlElement("w:sdtPr")
    tag_element = OxmlElement("w:tag")
    tag_element.set(qn("w:val"), tag)
    properties.append(tag_element)
    content = OxmlElement("w:sdtContent")
    instruction = f" REF {bookmark_name} \\n \\h " if heading_number_only else f" REF {bookmark_name} \\h "
    append_complex_field_to_container(
        content,
        instruction=instruction,
        cached_result=cached_number,
    )
    if alias is not None:
        alias_run = OxmlElement("w:r")
        alias_text = OxmlElement("w:t")
        alias_text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        alias_text.text = f" {alias}"
        alias_run.append(alias_text)
        content.append(alias_run)
    sdt.extend((properties, content))
    return sdt


def append_complex_field_to_container(
    container: Any,
    *,
    instruction: str,
    cached_result: str,
) -> None:
    """Append the same complex-field skeleton to an arbitrary run container."""

    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    begin_run = OxmlElement("w:r")
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    begin.set(qn("w:dirty"), "true")
    begin_run.append(begin)
    instruction_run = OxmlElement("w:r")
    instruction_text = OxmlElement("w:instrText")
    instruction_text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instruction_text.text = instruction
    instruction_run.append(instruction_text)
    separate_run = OxmlElement("w:r")
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    separate_run.append(separate)
    result_run = OxmlElement("w:r")
    result = OxmlElement("w:t")
    result.text = cached_result
    result_run.append(result)
    end_run = OxmlElement("w:r")
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    end_run.append(end)
    container.extend((begin_run, instruction_run, separate_run, result_run, end_run))


def wrap_direct_body_block(element: Any, tag: str) -> None:
    wrap_direct_body_group((element,), tag)


def wrap_direct_body_group(elements: tuple[Any, ...], tag: str) -> None:
    """Wrap one contiguous logical main-body block group in one outer SDT."""

    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    if not elements:
        raise DocxSemanticsV3Error("v3 block projection requires a non-empty logical group")
    parent = elements[0].getparent()
    if parent is None or parent.tag != qn("w:body"):
        raise DocxSemanticsV3Error("v3 block projection requires a direct main-body element")
    positions = []
    for element in elements:
        if element.getparent() is not parent:
            raise DocxSemanticsV3Error("v3 logical group must share one direct main-body parent")
        positions.append(parent.index(element))
    if positions != list(range(positions[0], positions[0] + len(elements))):
        raise DocxSemanticsV3Error("v3 logical group must be contiguous in the main body")
    index = positions[0]
    sdt = OxmlElement("w:sdt")
    properties = OxmlElement("w:sdtPr")
    tag_element = OxmlElement("w:tag")
    tag_element.set(qn("w:val"), tag)
    properties.append(tag_element)
    content = OxmlElement("w:sdtContent")
    for element in elements:
        parent.remove(element)
        content.append(element)
    sdt.extend((properties, content))
    parent.insert(index, sdt)


def sdt_tag(sdt: Any) -> str | None:
    from docx.oxml.ns import qn

    properties = sdt.find(qn("w:sdtPr"))
    if properties is None:
        return None
    tags = properties.findall(qn("w:tag"))
    if len(tags) != 1:
        return None
    return tags[0].get(qn("w:val"))


def visible_text(element: Any) -> str:
    from docx.oxml.ns import qn

    parts: list[str] = []
    for child in element.iter():
        if child.tag == qn("w:t"):
            parts.append(child.text or "")
        elif child.tag == qn("w:tab"):
            parts.append("\t")
        elif child.tag in {qn("w:br"), qn("w:cr")}:
            parts.append("\n")
    return "".join(parts)


def recover_paragraph_children(
    children: list[Any],
    *,
    target_ids_by_bookmark: dict[str, str],
    soft_tokens_by_tag: dict[str, str],
    occurrence_tokens_by_tag: dict[str, str] | None = None,
    stable_reference_target_ids: list[str] | None = None,
) -> tuple[str, bool]:
    from docx.oxml.ns import qn

    output: list[str] = []
    semantic = False
    index = 0
    while index < len(children):
        child = children[index]
        if child.tag == qn("w:sdt"):
            tag = sdt_tag(child)
            if tag in soft_tokens_by_tag:
                output.append(soft_tokens_by_tag[tag])
                semantic = True
            elif occurrence_tokens_by_tag is not None and tag in occurrence_tokens_by_tag:
                output.append(occurrence_tokens_by_tag[tag])
                semantic = True
                if stable_reference_target_ids is not None:
                    content = child.find(qn("w:sdtContent"))
                    instruction = "".join(
                        item.text or "" for item in (() if content is None else content.iter(qn("w:instrText")))
                    )
                    match = _REF_INSTRUCTION_RE.fullmatch(instruction)
                    if match is None or match.group(1) not in target_ids_by_bookmark:
                        raise DocxSemanticsV3Error("reference-occurrence REF field has no authenticated target")
                    stable_reference_target_ids.append(target_ids_by_bookmark[match.group(1)])
            else:
                output.append(visible_text(child))
            index += 1
            continue
        begin = child.find(f".//{qn('w:fldChar')}[@{qn('w:fldCharType')}='begin']")
        if child.tag == qn("w:r") and begin is not None:
            end_index = index
            instruction_parts: list[str] = []
            field_visible: list[str] = []
            after_separator = False
            while end_index < len(children):
                field_child = children[end_index]
                instruction_parts.extend(item.text or "" for item in field_child.iter(qn("w:instrText")))
                char_types = [item.get(qn("w:fldCharType")) for item in field_child.iter(qn("w:fldChar"))]
                if "separate" in char_types:
                    after_separator = True
                elif after_separator and "end" not in char_types:
                    field_visible.append(visible_text(field_child))
                if "end" in char_types:
                    break
                end_index += 1
            if end_index >= len(children):
                raise DocxSemanticsV3Error("unterminated complex field in semantic paragraph")
            instruction = "".join(instruction_parts)
            match = _REF_INSTRUCTION_RE.fullmatch(instruction)
            if match is not None:
                bookmark_name = match.group(1)
                if bookmark_name not in target_ids_by_bookmark:
                    raise DocxSemanticsV3Error("owned REF field has no authenticated semantic target")
                target_id = target_ids_by_bookmark[bookmark_name]
                output.append(f"@[[#^{target_id}]]")
                semantic = True
                if stable_reference_target_ids is not None:
                    stable_reference_target_ids.append(target_id)
            elif re.search(r"\bREF\s+DW_T_", instruction, re.IGNORECASE) is not None:
                raise DocxSemanticsV3Error("owned REF field instruction is not canonical")
            else:
                output.append("".join(field_visible))
            index = end_index + 1
            continue
        output.append(visible_text(child))
        index += 1
    return "".join(output), semantic


def soft_reference_visible_text(authored_token: str, cached_number: str) -> str:
    body = authored_token[3:-2]
    _selector, separator, alias = body.partition("|")
    return cached_number if not separator else f"{cached_number} {alias}"
