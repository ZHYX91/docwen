"""
DOCX 批注锚点分析报告生成器

解析 DOCX 的 word/document.xml 与 word/comments.xml，提取每个批注的锚点范围
（w:commentRangeStart / w:commentRangeEnd）以及对应的覆盖文本，并生成 Markdown 报告。

实现要点：
- 仅使用标准库（zipfile + xml.etree.ElementTree），不依赖 python-docx
- 支持可选脱敏输出，避免把敏感内容写入检查产物
"""

from __future__ import annotations

import datetime as _dt
import re
import zipfile
from collections.abc import Iterable
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_NS = {"w": _W_NS}


def _sort_comment_id(cid: str):
    return int(cid) if cid.isdigit() else cid


@dataclass(frozen=True)
class AnchorOccurrence:
    comment_id: str
    paragraph_index: int
    start: int
    end: int
    covered_text: str
    context_before: str
    context_after: str


@dataclass(frozen=True)
class CrossParagraphAnchor:
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
    cross_paragraph: list[CrossParagraphAnchor]
    end_without_start_ids: list[str]
    start_without_end_ids: list[str]


def read_docx_part(docx_path: str | Path, part_path: str) -> bytes | None:
    with zipfile.ZipFile(docx_path) as zf:
        try:
            return zf.read(part_path)
        except KeyError:
            return None


def _qn_attr(local: str) -> str:
    return f"{{{_W_NS}}}{local}"


def _get_w_id(elem: ET.Element) -> str | None:
    return elem.attrib.get(_qn_attr("id"))


def _iter_paragraph_events(p: ET.Element) -> Iterable[tuple[str, str | None]]:
    for elem in p.iter():
        tag = elem.tag
        if tag == _qn_attr("commentRangeStart"):
            yield ("start", _get_w_id(elem))
        elif tag == _qn_attr("commentRangeEnd"):
            yield ("end", _get_w_id(elem))
        elif tag == _qn_attr("t"):
            yield ("text", elem.text or "")
        elif tag == _qn_attr("tab"):
            yield ("text", "\t")
        elif tag == _qn_attr("noBreakHyphen"):
            yield ("text", "\u2011")
        elif tag in {_qn_attr("br"), _qn_attr("cr")}:
            yield ("text", "\n")
        elif tag == _qn_attr("sym"):
            sym = elem.attrib.get(_qn_attr("char"))
            if sym:
                try:
                    yield ("text", chr(int(sym, 16)))
                except ValueError:
                    yield ("text", "\ufffd")
            else:
                yield ("text", "\ufffd")


def extract_occurrences_from_document_xml(
    document_xml: bytes,
    context_chars: int,
    redact: bool,
) -> tuple[list[AnchorOccurrence], AnchorDiagnostics]:
    root = ET.fromstring(document_xml)
    paragraphs: list[str] = []
    events: list[tuple[str, str, int, int]] = []

    for para_index, p in enumerate(root.findall(".//w:p", _NS), start=1):
        buf: list[str] = []
        current_len = 0
        for kind, payload in _iter_paragraph_events(p):
            if kind in {"start", "end"}:
                if payload is None:
                    continue
                events.append((kind, payload, para_index, current_len))
                continue

            text = payload or ""
            buf.append(text)
            current_len += len(text)

        paragraphs.append("".join(buf))

    def _ctx(s: str, pos: int) -> tuple[str, str]:
        before = s[max(0, pos - context_chars) : pos]
        after = s[pos : pos + context_chars]
        if redact:
            before = redact_text(before)
            after = redact_text(after)
        return before, after

    occurrences: list[AnchorOccurrence] = []
    cross_paragraph: list[CrossParagraphAnchor] = []
    end_without_start: set[str] = set()
    open_ranges: dict[str, tuple[int, int]] = {}

    for kind, cid, para_index, offset in events:
        if kind == "start":
            open_ranges[cid] = (para_index, offset)
            continue

        start_info = open_ranges.pop(cid, None)
        if start_info is None:
            end_without_start.add(cid)
            continue

        start_para, start_off = start_info
        if start_para != para_index:
            start_text = paragraphs[start_para - 1] if 1 <= start_para <= len(paragraphs) else ""
            end_text = paragraphs[para_index - 1] if 1 <= para_index <= len(paragraphs) else ""
            s_before, s_after = _ctx(start_text, start_off)
            e_before, e_after = _ctx(end_text, offset)
            cross_paragraph.append(
                CrossParagraphAnchor(
                    comment_id=cid,
                    start_paragraph_index=start_para,
                    start_offset=start_off,
                    end_paragraph_index=para_index,
                    end_offset=offset,
                    start_context_before=s_before,
                    start_context_after=s_after,
                    end_context_before=e_before,
                    end_context_after=e_after,
                )
            )
            continue

        para_text = paragraphs[para_index - 1] if 1 <= para_index <= len(paragraphs) else ""
        covered = para_text[start_off:offset]
        before = para_text[max(0, start_off - context_chars) : start_off]
        after = para_text[offset : offset + context_chars]
        if redact:
            covered = redact_text(covered)
            before = redact_text(before)
            after = redact_text(after)

        occurrences.append(
            AnchorOccurrence(
                comment_id=cid,
                paragraph_index=para_index,
                start=start_off,
                end=offset,
                covered_text=covered,
                context_before=before,
                context_after=after,
            )
        )

    start_without_end_ids = sorted(open_ranges.keys(), key=_sort_comment_id)
    end_without_start_ids = sorted(end_without_start, key=_sort_comment_id)
    cross_paragraph_sorted = sorted(cross_paragraph, key=lambda x: _sort_comment_id(x.comment_id))
    diagnostics = AnchorDiagnostics(
        cross_paragraph=cross_paragraph_sorted,
        end_without_start_ids=end_without_start_ids,
        start_without_end_ids=start_without_end_ids,
    )
    return occurrences, diagnostics


def extract_comment_texts_from_comments_xml(comments_xml: bytes, redact: bool) -> dict[str, str]:
    root = ET.fromstring(comments_xml)
    out: dict[str, str] = {}
    for comment in root.findall(".//w:comment", _NS):
        cid = _get_w_id(comment)
        if cid is None:
            continue
        parts: list[str] = []
        for t in comment.findall(".//w:t", _NS):
            parts.append(t.text or "")
        text = "".join(parts).strip()
        out[cid] = redact_text(text) if redact else text
    return out


_RE_REDACT = re.compile(
    r"["
    r"0-9A-Za-z"  # ASCII 字母数字
    r"\u00C0-\u024F"  # 拉丁扩展（含重音：àéîöü 等，覆盖法/德/西/葡）
    r"\u0400-\u04FF"  # 西里尔字母（俄文）
    r"\u1100-\u11FF"  # 韩文字母（Jamo）
    r"\u1E00-\u1EFF"  # 拉丁扩展（越南文等）
    r"\u3040-\u309F"  # 日文平假名
    r"\u30A0-\u30FF"  # 日文片假名
    r"\u3400-\u4DBF"  # CJK 扩展 A
    r"\u4E00-\u9FFF"  # CJK 基本区
    r"\uAC00-\uD7AF"  # 韩文音节
    r"\uF900-\uFAFF"  # CJK 兼容汉字
    r"\uFF10-\uFF19"  # 全角数字
    r"\uFF21-\uFF3A"  # 全角大写字母
    r"\uFF41-\uFF5A"  # 全角小写字母
    r"]"
)


def redact_text(s: str) -> str:
    """将文本中的字母、数字、汉字等可辨识字符替换为 █，保留标点和符号结构。"""
    return _RE_REDACT.sub("█", s)


def build_anchor_report_markdown(
    docx_path: str | Path,
    *,
    context_chars: int = 20,
    redact: bool = False,
) -> str:
    docx_path = str(docx_path)
    document_xml = read_docx_part(docx_path, "word/document.xml")
    if document_xml is None:
        raise FileNotFoundError("DOCX 缺少 word/document.xml")

    comments_xml = read_docx_part(docx_path, "word/comments.xml")
    occurrences, diagnostics = extract_occurrences_from_document_xml(document_xml, context_chars, redact)
    comments = extract_comment_texts_from_comments_xml(comments_xml, redact) if comments_xml else {}

    ids_in_occ = {o.comment_id for o in occurrences}
    ids_in_comments = set(comments.keys())
    ids_missing_ranges = sorted(ids_in_comments - ids_in_occ, key=_sort_comment_id)
    ids_missing_comments = sorted(ids_in_occ - ids_in_comments, key=_sort_comment_id)

    lines: list[str] = []
    lines.append("# DOCX 批注锚点报告")
    lines.append("")
    lines.append(f"- 输入文件：`{Path(docx_path).name}`")
    lines.append(f"- 生成时间：`{_dt.datetime.now().isoformat(timespec='seconds')}`")
    lines.append(f"- 批注数（comments.xml）：`{len(comments)}`")
    lines.append(f"- 锚点范围数（document.xml）：`{len(occurrences)}`")
    lines.append("")

    if diagnostics.cross_paragraph or diagnostics.end_without_start_ids or diagnostics.start_without_end_ids:
        lines.append("## 锚点异常（跨段/未闭合）")
        lines.append("")
        if diagnostics.cross_paragraph:
            lines.append("### 跨段锚点")
            lines.append("")
            for item in diagnostics.cross_paragraph:
                lines.append(
                    f"- `{item.comment_id}`：段落 `{item.start_paragraph_index}` `@{item.start_offset}`"
                    f" → 段落 `{item.end_paragraph_index}` `@{item.end_offset}`"
                )
                lines.append("")
                lines.append("```text")
                lines.append(f"START: {item.start_context_before}<<START>>{item.start_context_after}")
                lines.append(f"END  : {item.end_context_before}<<END>>{item.end_context_after}")
                lines.append("```")
                lines.append("")
        if diagnostics.end_without_start_ids:
            lines.append("### 只有 End 没有 Start 的 ID")
            lines.append("")
            for cid in diagnostics.end_without_start_ids:
                lines.append(f"- `{cid}`")
            lines.append("")
        if diagnostics.start_without_end_ids:
            lines.append("### 只有 Start 没有 End 的 ID")
            lines.append("")
            for cid in diagnostics.start_without_end_ids:
                lines.append(f"- `{cid}`")
            lines.append("")

    if ids_missing_ranges:
        lines.append("## 未找到锚点范围的批注 ID")
        lines.append("")
        for cid in ids_missing_ranges:
            lines.append(f"- `{cid}`")
        lines.append("")

    if ids_missing_comments:
        lines.append("## 未找到批注文案的锚点 ID")
        lines.append("")
        for cid in ids_missing_comments:
            lines.append(f"- `{cid}`")
        lines.append("")

    by_id: dict[str, list[AnchorOccurrence]] = {}
    for occ in occurrences:
        by_id.setdefault(occ.comment_id, []).append(occ)

    for cid in sorted(by_id.keys(), key=_sort_comment_id):
        lines.append(f"## 批注 `{cid}`")
        lines.append("")
        comment_text = comments.get(cid, "")
        if comment_text:
            lines.append(f"- 批注内容：{comment_text}")
        lines.append(f"- 锚点次数：`{len(by_id[cid])}`")
        lines.append("")
        for occ in by_id[cid]:
            lines.append(
                f"- 段落：`{occ.paragraph_index}`，范围：`[{occ.start},{occ.end})`，"
                f"覆盖：`{len(occ.covered_text)}` 字符"
            )
            lines.append("")
            lines.append("```text")
            lines.append(f"{occ.context_before}[{occ.covered_text}]{occ.context_after}")
            lines.append("```")
            lines.append("")

    return "\n".join(lines).rstrip() + "\n"
