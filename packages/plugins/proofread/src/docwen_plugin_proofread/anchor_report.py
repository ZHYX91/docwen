"""Analyze comment anchor ranges in a proofread DOCX and render Markdown.

The report is intentionally independent of ``python-docx``.  It reads the
Open XML parts directly so run-split anchors, cross-paragraph ranges, and
malformed/unclosed ranges remain visible in release evidence.  Optional
redaction preserves punctuation and report structure while masking readable
letters and digits.
"""

from __future__ import annotations

import hashlib
import html
import io
import unicodedata
import zipfile
from collections.abc import Iterable
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_NS = {"w": _W_NS}
_MAX_DOCX_BYTES = 256 * 1024 * 1024
_MAX_XML_PART_BYTES = 64 * 1024 * 1024
_MAX_COMPRESSION_RATIO = 200
_MAX_CONTEXT_CHARS = 4096


def _sort_comment_id(comment_id: str) -> tuple[int, int | str]:
    """Sort numeric comment IDs numerically, then non-numeric IDs lexically."""
    if comment_id.isdigit():
        return (0, int(comment_id))
    return (1, comment_id)


@dataclass(frozen=True)
class CommentAnchorInfo:
    """Compatibility projection used by current proofread parity probes."""

    comment_id: str
    author: str
    text: str
    anchor_paragraph_index: int | None


@dataclass(frozen=True)
class AnchorOccurrence:
    """One complete comment range contained within a single paragraph."""

    comment_id: str
    paragraph_index: int
    start: int
    end: int
    covered_text: str
    context_before: str
    context_after: str


@dataclass(frozen=True)
class CrossParagraphAnchor:
    """A complete comment range whose endpoints are in different paragraphs."""

    comment_id: str
    start_paragraph_index: int
    start_offset: int
    end_paragraph_index: int
    end_offset: int
    start_context_before: str
    start_context_after: str
    end_context_before: str
    end_context_after: str


@dataclass(frozen=True)
class AnchorDiagnostics:
    """Structural range problems found while walking ``document.xml``."""

    cross_paragraph: list[CrossParagraphAnchor]
    end_without_start_ids: list[str]
    start_without_end_ids: list[str]


def read_docx_part(docx_path: str | Path, part_path: str) -> bytes | None:
    """Read one DOCX ZIP part, returning ``None`` when the part is absent."""
    with zipfile.ZipFile(docx_path) as archive:
        try:
            return archive.read(part_path)
        except KeyError:
            return None


def _validate_context_chars(context_chars: int) -> None:
    if isinstance(context_chars, bool) or not isinstance(context_chars, int):
        raise TypeError("context_chars must be an integer")
    if not 0 <= context_chars <= _MAX_CONTEXT_CHARS:
        raise ValueError(f"context_chars must be between 0 and {_MAX_CONTEXT_CHARS}")


def _validate_xml_bytes(xml: bytes, part_path: str) -> None:
    if len(xml) > _MAX_XML_PART_BYTES:
        raise ValueError(f"DOCX XML part is too large: {part_path}")
    lowered = xml.lower()
    if b"<!doctype" in lowered or b"<!entity" in lowered:
        raise ValueError(f"DTD/entity declarations are not allowed in {part_path}")


def _read_anchor_snapshot(path: Path) -> tuple[bytes, bytes | None, str]:
    """Read both anchor parts from one immutable in-memory DOCX snapshot."""
    payload = path.read_bytes()
    if len(payload) > _MAX_DOCX_BYTES:
        raise ValueError("DOCX is too large for anchor reporting")
    source_sha256 = hashlib.sha256(payload).hexdigest()

    with zipfile.ZipFile(io.BytesIO(payload)) as archive:
        parts: dict[str, bytes | None] = {}
        for part_path, required in (("word/document.xml", True), ("word/comments.xml", False)):
            infos = [info for info in archive.infolist() if info.filename == part_path]
            if len(infos) > 1:
                raise ValueError(f"DOCX contains duplicate ZIP member: {part_path}")
            if not infos:
                if required:
                    raise FileNotFoundError(f"DOCX 缺少 {part_path}")
                parts[part_path] = None
                continue

            info = infos[0]
            if info.flag_bits & 0x1:
                raise ValueError(f"Encrypted DOCX ZIP member is not supported: {part_path}")
            if info.file_size > _MAX_XML_PART_BYTES:
                raise ValueError(f"DOCX XML part is too large: {part_path}")
            if (
                info.file_size > 1024 * 1024
                and info.compress_size > 0
                and info.file_size / info.compress_size > _MAX_COMPRESSION_RATIO
            ):
                raise ValueError(f"DOCX XML compression ratio is unsafe: {part_path}")
            xml = archive.read(info)
            _validate_xml_bytes(xml, part_path)
            parts[part_path] = xml

    document_xml = parts["word/document.xml"]
    assert document_xml is not None
    return document_xml, parts["word/comments.xml"], source_sha256


def _qn_attr(local: str) -> str:
    return f"{{{_W_NS}}}{local}"


def _get_w_id(element: ET.Element) -> str | None:
    return element.attrib.get(_qn_attr("id"))


def _iter_paragraph_events(paragraph: ET.Element) -> Iterable[tuple[str, str | None]]:
    """Yield range-boundary and visible-text events in XML document order."""
    for element in paragraph.iter():
        tag = element.tag
        if tag == _qn_attr("commentRangeStart"):
            yield ("start", _get_w_id(element))
        elif tag == _qn_attr("commentRangeEnd"):
            yield ("end", _get_w_id(element))
        elif tag == _qn_attr("t"):
            yield ("text", element.text or "")
        elif tag == _qn_attr("tab"):
            yield ("text", "\t")
        elif tag == _qn_attr("noBreakHyphen"):
            yield ("text", "\u2011")
        elif tag in {_qn_attr("br"), _qn_attr("cr")}:
            yield ("text", "\n")
        elif tag == _qn_attr("sym"):
            symbol = element.attrib.get(_qn_attr("char"))
            if symbol:
                try:
                    yield ("text", chr(int(symbol, 16)))
                except ValueError:
                    yield ("text", "\ufffd")
            else:
                yield ("text", "\ufffd")


def extract_occurrences_from_document_xml(
    document_xml: bytes,
    context_chars: int,
    redact: bool,
) -> tuple[list[AnchorOccurrence], AnchorDiagnostics]:
    """Extract complete ranges plus cross-paragraph and unclosed diagnostics."""
    _validate_context_chars(context_chars)
    _validate_xml_bytes(document_xml, "word/document.xml")
    root = ET.fromstring(document_xml)
    paragraphs: list[str] = []
    events: list[tuple[str, str, int, int]] = []

    for paragraph_index, paragraph in enumerate(root.findall(".//w:p", _NS), start=1):
        buffer: list[str] = []
        current_length = 0
        for kind, payload in _iter_paragraph_events(paragraph):
            if kind in {"start", "end"}:
                if payload is not None:
                    events.append((kind, payload, paragraph_index, current_length))
                continue

            text = payload or ""
            buffer.append(text)
            current_length += len(text)

        paragraphs.append("".join(buffer))

    def context(text: str, position: int) -> tuple[str, str]:
        before = text[max(0, position - context_chars) : position]
        after = text[position : position + context_chars]
        if redact:
            before = redact_text(before)
            after = redact_text(after)
        return before, after

    occurrences: list[AnchorOccurrence] = []
    cross_paragraph: list[CrossParagraphAnchor] = []
    end_without_start: set[str] = set()
    open_ranges: dict[str, tuple[int, int]] = {}

    for kind, comment_id, paragraph_index, offset in events:
        if kind == "start":
            open_ranges[comment_id] = (paragraph_index, offset)
            continue

        start_info = open_ranges.pop(comment_id, None)
        if start_info is None:
            end_without_start.add(comment_id)
            continue

        start_paragraph, start_offset = start_info
        if start_paragraph != paragraph_index:
            start_text = paragraphs[start_paragraph - 1] if 1 <= start_paragraph <= len(paragraphs) else ""
            end_text = paragraphs[paragraph_index - 1] if 1 <= paragraph_index <= len(paragraphs) else ""
            start_before, start_after = context(start_text, start_offset)
            end_before, end_after = context(end_text, offset)
            cross_paragraph.append(
                CrossParagraphAnchor(
                    comment_id=comment_id,
                    start_paragraph_index=start_paragraph,
                    start_offset=start_offset,
                    end_paragraph_index=paragraph_index,
                    end_offset=offset,
                    start_context_before=start_before,
                    start_context_after=start_after,
                    end_context_before=end_before,
                    end_context_after=end_after,
                )
            )
            continue

        paragraph_text = paragraphs[paragraph_index - 1] if 1 <= paragraph_index <= len(paragraphs) else ""
        covered = paragraph_text[start_offset:offset]
        before = paragraph_text[max(0, start_offset - context_chars) : start_offset]
        after = paragraph_text[offset : offset + context_chars]
        if redact:
            covered = redact_text(covered)
            before = redact_text(before)
            after = redact_text(after)

        occurrences.append(
            AnchorOccurrence(
                comment_id=comment_id,
                paragraph_index=paragraph_index,
                start=start_offset,
                end=offset,
                covered_text=covered,
                context_before=before,
                context_after=after,
            )
        )

    diagnostics = AnchorDiagnostics(
        cross_paragraph=sorted(cross_paragraph, key=lambda item: _sort_comment_id(item.comment_id)),
        end_without_start_ids=sorted(end_without_start, key=_sort_comment_id),
        start_without_end_ids=sorted(open_ranges, key=_sort_comment_id),
    )
    return occurrences, diagnostics


def extract_comment_texts_from_comments_xml(comments_xml: bytes, redact: bool) -> dict[str, str]:
    """Map comment IDs to concatenated comment text."""
    _validate_xml_bytes(comments_xml, "word/comments.xml")
    root = ET.fromstring(comments_xml)
    output: dict[str, str] = {}
    for comment in root.findall(".//w:comment", _NS):
        comment_id = _get_w_id(comment)
        if comment_id is None:
            continue
        text = "".join(node.text or "" for node in comment.findall(".//w:t", _NS)).strip()
        output[comment_id] = redact_text(text) if redact else text
    return output


def redact_text(text: str) -> str:
    """Mask readable multilingual characters while preserving punctuation."""
    return "".join("█" if unicodedata.category(character)[0] in {"L", "M", "N"} else character for character in text)


def _longest_backtick_run(text: str) -> int:
    longest = current = 0
    for character in text:
        if character == "`":
            current += 1
            longest = max(longest, current)
        else:
            current = 0
    return longest


def _inline_code(value: object) -> str:
    text = str(value).replace("\r", " ").replace("\n", " ")
    fence = "`" * max(1, _longest_backtick_run(text) + 1)
    padding = " " if text.startswith(("`", " ")) or text.endswith(("`", " ")) else ""
    return f"{fence}{padding}{text}{padding}{fence}"


def _escaped_markdown_text(value: object) -> str:
    text = str(value).replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
    text = text.replace("\\", "\\\\")
    text = html.escape(text, quote=False)
    for character in "`*_{}[]()#+-.!|>":
        text = text.replace(character, f"\\{character}")
    return text


def _fenced_text(*lines: str) -> list[str]:
    fence = "`" * max(3, _longest_backtick_run("\n".join(lines)) + 1)
    return [f"{fence}text", *lines, fence]


def build_anchor_report_markdown(
    docx_path: str | Path,
    *,
    context_chars: int = 20,
    redact: bool = False,
) -> str:
    """Build the complete historical DOCX comment-anchor Markdown report."""
    _validate_context_chars(context_chars)
    path = Path(docx_path)
    document_xml, comments_xml, source_sha256 = _read_anchor_snapshot(path)
    occurrences, diagnostics = extract_occurrences_from_document_xml(document_xml, context_chars, redact)
    comments = extract_comment_texts_from_comments_xml(comments_xml, redact) if comments_xml else {}

    ids_with_ranges = {occurrence.comment_id for occurrence in occurrences}
    ids_with_ranges.update(item.comment_id for item in diagnostics.cross_paragraph)
    ids_in_comments = set(comments)
    ids_missing_ranges = sorted(ids_in_comments - ids_with_ranges, key=_sort_comment_id)
    ids_missing_comments = sorted(ids_with_ranges - ids_in_comments, key=_sort_comment_id)

    lines = [
        "# DOCX 批注锚点报告",
        "",
        f"- 输入文件：{_inline_code(redact_text(path.name) if redact else path.name)}",
        f"- 输入 SHA-256：{_inline_code(source_sha256)}",
        f"- 批注数（comments.xml）：{_inline_code(len(comments))}",
        f"- 锚点范围数（document.xml）：{_inline_code(len(occurrences))}",
        "",
    ]

    if diagnostics.cross_paragraph or diagnostics.end_without_start_ids or diagnostics.start_without_end_ids:
        lines.extend(["## 锚点异常（跨段/未闭合）", ""])
        if diagnostics.cross_paragraph:
            lines.extend(["### 跨段锚点", ""])
            for item in diagnostics.cross_paragraph:
                lines.extend(
                    [
                        f"- {_inline_code(item.comment_id)}：段落 {_inline_code(item.start_paragraph_index)} "
                        f"{_inline_code(f'@{item.start_offset}')} → 段落 {_inline_code(item.end_paragraph_index)} "
                        f"{_inline_code(f'@{item.end_offset}')}",
                        "",
                        *_fenced_text(
                            f"START: {item.start_context_before}<<START>>{item.start_context_after}",
                            f"END  : {item.end_context_before}<<END>>{item.end_context_after}",
                        ),
                        "",
                    ]
                )
        if diagnostics.end_without_start_ids:
            lines.extend(["### 只有 End 没有 Start 的 ID", ""])
            lines.extend(f"- {_inline_code(comment_id)}" for comment_id in diagnostics.end_without_start_ids)
            lines.append("")
        if diagnostics.start_without_end_ids:
            lines.extend(["### 只有 Start 没有 End 的 ID", ""])
            lines.extend(f"- {_inline_code(comment_id)}" for comment_id in diagnostics.start_without_end_ids)
            lines.append("")

    if ids_missing_ranges:
        lines.extend(["## 未找到锚点范围的批注 ID", ""])
        lines.extend(f"- {_inline_code(comment_id)}" for comment_id in ids_missing_ranges)
        lines.append("")

    if ids_missing_comments:
        lines.extend(["## 未找到批注文案的锚点 ID", ""])
        lines.extend(f"- {_inline_code(comment_id)}" for comment_id in ids_missing_comments)
        lines.append("")

    by_id: dict[str, list[AnchorOccurrence]] = {}
    for occurrence in occurrences:
        by_id.setdefault(occurrence.comment_id, []).append(occurrence)

    for comment_id in sorted(by_id, key=_sort_comment_id):
        lines.extend([f"## 批注 {_inline_code(comment_id)}", ""])
        comment_text = comments.get(comment_id, "")
        if comment_text:
            lines.append(f"- 批注内容：{_escaped_markdown_text(comment_text)}")
        lines.extend([f"- 锚点次数：{_inline_code(len(by_id[comment_id]))}", ""])
        for occurrence in by_id[comment_id]:
            lines.extend(
                [
                    f"- 段落：{_inline_code(occurrence.paragraph_index)}，"
                    f"范围：{_inline_code(f'[{occurrence.start},{occurrence.end})')}，"
                    f"覆盖：{_inline_code(len(occurrence.covered_text))} 字符",
                    "",
                    *_fenced_text(f"{occurrence.context_before}[{occurrence.covered_text}]{occurrence.context_after}"),
                    "",
                ]
            )

    return "\n".join(lines).rstrip() + "\n"


def _extract_comments(docx_path: Path) -> list[CommentAnchorInfo]:
    """Return the current compatibility projection used by parity probes.

    Paragraph indices intentionally remain zero-based to preserve the API that
    was already consumed by the current proofread tests and validation probe.
    """
    with zipfile.ZipFile(docx_path, "r") as archive:
        if "word/comments.xml" not in archive.namelist():
            return []
        comments_tree = ET.fromstring(archive.read("word/comments.xml"))
        paragraph_comments: dict[str, int] = {}
        if "word/document.xml" in archive.namelist():
            document_tree = ET.fromstring(archive.read("word/document.xml"))
            _map_comment_refs(document_tree, paragraph_comments)

    comments: list[CommentAnchorInfo] = []
    for element in comments_tree.iter(_qn_attr("comment")):
        comment_id = element.get(_qn_attr("id"), "")
        author = element.get(_qn_attr("author"), "unknown")
        text = "".join(node.text or "" for node in element.iter(_qn_attr("t")))
        comments.append(
            CommentAnchorInfo(
                comment_id=comment_id,
                author=author,
                text=text,
                anchor_paragraph_index=paragraph_comments.get(comment_id),
            )
        )
    return comments


def _map_comment_refs(document_tree: ET.Element, output: dict[str, int]) -> None:
    """Map range-start IDs in direct body paragraphs to zero-based indices."""
    body = document_tree.find(f".//{_qn_attr('body')}")
    if body is None:
        return

    paragraph_index = 0
    for child in list(body):
        if child.tag != _qn_attr("p"):
            continue
        for reference in child.iter(_qn_attr("commentRangeStart")):
            comment_id = reference.get(_qn_attr("id"), "")
            if comment_id and comment_id not in output:
                output[comment_id] = paragraph_index
        paragraph_index += 1
