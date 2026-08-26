from __future__ import annotations

import inspect
import zipfile
from pathlib import Path

import pytest

import docwen_plugin_proofread.anchor_report as anchor_report
from docwen_plugin_proofread.anchor_report import (
    CommentAnchorInfo,
    _extract_comments,
    build_anchor_report_markdown,
    extract_comment_texts_from_comments_xml,
    extract_occurrences_from_document_xml,
    redact_text,
)

pytestmark = pytest.mark.unit

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _document_xml(body: str) -> bytes:
    return (
        f'<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="{_W_NS}"><w:body>{body}</w:body></w:document>'
    ).encode()


def _comments_xml(*comments: tuple[str, str, str]) -> bytes:
    content = "".join(
        f'<w:comment w:id="{comment_id}" w:author="{author}"><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:comment>'
        for comment_id, author, text in comments
    )
    return f'<w:comments xmlns:w="{_W_NS}">{content}</w:comments>'.encode()


def _write_docx(
    tmp_path: Path,
    *,
    document_xml: bytes | None,
    comments_xml: bytes | None,
) -> Path:
    output = tmp_path / "input.docx"
    with zipfile.ZipFile(output, "w") as archive:
        if document_xml is not None:
            archive.writestr("word/document.xml", document_xml)
        if comments_xml is not None:
            archive.writestr("word/comments.xml", comments_xml)
    return output


def test_public_signature_keeps_historical_context_and_redaction_defaults() -> None:
    signature = inspect.signature(build_anchor_report_markdown)

    assert signature.parameters["context_chars"].default == 20
    assert signature.parameters["redact"].default is False


def test_report_is_byte_stable_and_contains_no_wall_clock(tmp_path: Path) -> None:
    document_xml = _document_xml("<w:p><w:r><w:t>稳定</w:t></w:r></w:p>")
    docx = _write_docx(tmp_path, document_xml=document_xml, comments_xml=None)

    first = build_anchor_report_markdown(docx)
    second = build_anchor_report_markdown(docx)

    assert first == second
    assert "生成时间" not in first
    assert "datetime" not in inspect.getsource(anchor_report.build_anchor_report_markdown)
    assert "输入 SHA-256" in first


@pytest.mark.parametrize("context_chars", [-1, 4097, True])
def test_context_chars_is_a_bounded_non_boolean_integer(tmp_path: Path, context_chars: int) -> None:
    document_xml = _document_xml("<w:p><w:r><w:t>边界</w:t></w:r></w:p>")
    docx = _write_docx(tmp_path, document_xml=document_xml, comments_xml=None)

    with pytest.raises((TypeError, ValueError), match="context_chars"):
        build_anchor_report_markdown(docx, context_chars=context_chars)


def test_same_paragraph_range_keeps_half_open_offsets_covered_text_and_context() -> None:
    document_xml = _document_xml(
        '<w:p><w:r><w:t>前</w:t></w:r><w:commentRangeStart w:id="7"/>'
        '<w:r><w:t>错</w:t><w:tab/></w:r><w:commentRangeEnd w:id="7"/>'
        "<w:r><w:t>后</w:t></w:r></w:p>"
    )

    occurrences, diagnostics = extract_occurrences_from_document_xml(document_xml, 1, False)

    assert diagnostics.cross_paragraph == []
    assert diagnostics.end_without_start_ids == []
    assert diagnostics.start_without_end_ids == []
    assert len(occurrences) == 1
    occurrence = occurrences[0]
    assert (occurrence.comment_id, occurrence.paragraph_index, occurrence.start, occurrence.end) == ("7", 1, 1, 3)
    assert occurrence.covered_text == "错\t"
    assert occurrence.context_before == "前"
    assert occurrence.context_after == "后"


def test_cross_paragraph_and_unclosed_boundaries_are_reported_with_coordinates() -> None:
    document_xml = _document_xml(
        '<w:p><w:r><w:t>甲</w:t></w:r><w:commentRangeStart w:id="4"/></w:p>'
        '<w:p><w:r><w:t>乙</w:t></w:r><w:commentRangeEnd w:id="4"/>'
        '<w:commentRangeEnd w:id="2"/><w:commentRangeStart w:id="3"/></w:p>'
    )

    occurrences, diagnostics = extract_occurrences_from_document_xml(document_xml, 2, False)

    assert occurrences == []
    assert diagnostics.end_without_start_ids == ["2"]
    assert diagnostics.start_without_end_ids == ["3"]
    assert len(diagnostics.cross_paragraph) == 1
    cross = diagnostics.cross_paragraph[0]
    assert (
        cross.comment_id,
        cross.start_paragraph_index,
        cross.start_offset,
        cross.end_paragraph_index,
        cross.end_offset,
    ) == ("4", 1, 1, 2, 1)
    assert cross.start_context_before == "甲"
    assert cross.end_context_before == "乙"


def test_cross_paragraph_range_with_comment_is_not_reported_as_missing_range(tmp_path: Path) -> None:
    document_xml = _document_xml(
        '<w:p><w:commentRangeStart w:id="4"/><w:r><w:t>跨</w:t></w:r></w:p>'
        '<w:p><w:r><w:t>段</w:t></w:r><w:commentRangeEnd w:id="4"/></w:p>'
    )
    docx = _write_docx(
        tmp_path,
        document_xml=document_xml,
        comments_xml=_comments_xml(("4", "DocWen", "跨段批注")),
    )

    report = build_anchor_report_markdown(docx)

    assert "跨段锚点" in report
    assert "未找到锚点范围的批注 ID" not in report
    assert "未找到批注文案的锚点 ID" not in report


def test_cross_paragraph_range_without_comment_is_reported_as_missing_comment(tmp_path: Path) -> None:
    document_xml = _document_xml(
        '<w:p><w:commentRangeStart w:id="4"/><w:r><w:t>跨</w:t></w:r></w:p>'
        '<w:p><w:r><w:t>段</w:t></w:r><w:commentRangeEnd w:id="4"/></w:p>'
    )
    docx = _write_docx(tmp_path, document_xml=document_xml, comments_xml=None)

    report = build_anchor_report_markdown(docx)

    assert "未找到批注文案的锚点 ID" in report
    assert "- `4`" in report


def test_report_projects_ranges_comments_and_all_structural_diagnostics(tmp_path: Path) -> None:
    document_xml = _document_xml(
        '<w:p><w:r><w:t>前</w:t></w:r><w:commentRangeStart w:id="0"/>'
        '<w:r><w:t>错</w:t></w:r><w:commentRangeEnd w:id="0"/>'
        '<w:commentRangeEnd w:id="2"/><w:commentRangeStart w:id="3"/></w:p>'
        '<w:p><w:commentRangeStart w:id="4"/><w:r><w:t>跨</w:t></w:r></w:p>'
        '<w:p><w:r><w:t>段</w:t></w:r><w:commentRangeEnd w:id="4"/></w:p>'
    )
    docx = _write_docx(
        tmp_path,
        document_xml=document_xml,
        comments_xml=_comments_xml(("0", "DocWen", "测试批注"), ("9", "DocWen", "无范围")),
    )

    report = build_anchor_report_markdown(docx, context_chars=10, redact=False)

    assert "批注 `0`" in report
    assert "批注内容：测试批注" in report
    assert "范围：`[1,2)`" in report
    assert "前[错]" in report
    assert "跨段锚点" in report
    assert "`4`：段落 `2` `@0` → 段落 `3` `@1`" in report
    assert "只有 End 没有 Start 的 ID" in report and "- `2`" in report
    assert "只有 Start 没有 End 的 ID" in report and "- `3`" in report
    assert "未找到锚点范围的批注 ID" in report and "- `9`" in report


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("Hello123", "████████"),
        ("你好世界", "████"),
        ("あいうアイウ", "██████"),
        ("한국어", "███"),
        ("Привет", "██████"),
        ("àéîöüñ", "██████"),
        ("ếồắ", "███"),
        ("Ａ１ｂ２", "████"),
        ("（）、。！？「」", "（）、。！？「」"),
        ("张三（test）", "██（████）"),
    ],
)
def test_redact_text_masks_multilingual_content(text: str, expected: str) -> None:
    assert redact_text(text) == expected


def test_redact_text_masks_unicode_letters_marks_and_numbers_beyond_legacy_ranges() -> None:
    text = "Ωمرحباשלוםनमस्तेไทย\u0301𠀀１２३"

    redacted = redact_text(text)

    assert len(redacted) == len(text)
    assert set(redacted) == {"█"}


def test_redaction_covers_anchor_context_covered_text_and_comment_text(tmp_path: Path) -> None:
    document_xml = _document_xml(
        '<w:p><w:r><w:t>ABC中文</w:t></w:r><w:commentRangeStart w:id="1"/>'
        '<w:r><w:t>秘密</w:t></w:r><w:commentRangeEnd w:id="1"/></w:p>'
    )
    comments_xml = _comments_xml(("1", "Author", "敏感内容ABC123"))
    docx = _write_docx(tmp_path, document_xml=document_xml, comments_xml=comments_xml)

    report = build_anchor_report_markdown(docx, context_chars=20, redact=True)
    comments = extract_comment_texts_from_comments_xml(comments_xml, True)

    assert "ABC" not in report
    assert "中文" not in report
    assert "秘密" not in report
    assert "敏感内容" not in report
    assert comments == {"1": "██████████"}
    assert "[██]" in report


def test_untrusted_markdown_uses_dynamic_fences_and_escaped_comment_text(tmp_path: Path) -> None:
    document_xml = _document_xml(
        '<w:p><w:commentRangeStart w:id="x`x"/><w:r><w:t>```payload</w:t></w:r><w:commentRangeEnd w:id="x`x"/></w:p>'
    )
    comments_xml = _comments_xml(("x`x", "DocWen", "![x](https://example.invalid) &lt;img src=x&gt; **bold**"))
    docx = _write_docx(tmp_path, document_xml=document_xml, comments_xml=comments_xml)

    report = build_anchor_report_markdown(docx)

    assert "## 批注 ``x`x``" in report
    assert "````text\n[```payload]\n````" in report
    assert r"\!\[x\]\(https://example\.invalid\)" in report
    assert "&lt;img src=x&gt;" in report
    assert "<img src=x>" not in report
    assert r"\*\*bold\*\*" in report


def test_missing_comments_and_missing_ranges_are_distinguished(tmp_path: Path) -> None:
    range_only = _document_xml(
        '<w:p><w:commentRangeStart w:id="2"/><w:r><w:t>错</w:t></w:r><w:commentRangeEnd w:id="2"/></w:p>'
    )
    without_comments = _write_docx(tmp_path, document_xml=range_only, comments_xml=None)

    report = build_anchor_report_markdown(without_comments)

    assert "批注数（comments.xml）：`0`" in report
    assert "未找到批注文案的锚点 ID" in report
    assert "- `2`" in report


def test_missing_document_part_fails_closed(tmp_path: Path) -> None:
    docx = _write_docx(tmp_path, document_xml=None, comments_xml=_comments_xml(("1", "A", "B")))

    with pytest.raises(FileNotFoundError, match=r"word/document\.xml"):
        build_anchor_report_markdown(docx)


def test_duplicate_critical_zip_member_fails_closed(tmp_path: Path) -> None:
    docx = tmp_path / "duplicate.docx"
    document_xml = _document_xml("<w:p><w:r><w:t>重复</w:t></w:r></w:p>")
    with pytest.warns(UserWarning, match="Duplicate name"), zipfile.ZipFile(docx, "w") as archive:
        archive.writestr("word/document.xml", document_xml)
        archive.writestr("word/document.xml", document_xml)

    with pytest.raises(ValueError, match="duplicate ZIP member"):
        build_anchor_report_markdown(docx)


def test_dtd_or_entity_declarations_fail_closed(tmp_path: Path) -> None:
    unsafe_xml = (
        f'<!DOCTYPE w:document [<!ENTITY x "expanded">]><w:document xmlns:w="{_W_NS}">'
        "<w:body><w:p><w:r><w:t>&x;</w:t></w:r></w:p></w:body></w:document>"
    ).encode()
    docx = _write_docx(tmp_path, document_xml=unsafe_xml, comments_xml=None)

    with pytest.raises(ValueError, match="DTD/entity"):
        build_anchor_report_markdown(docx)


def test_extract_comments_preserves_zero_based_current_probe_projection(tmp_path: Path) -> None:
    document_xml = _document_xml(
        "<w:p><w:r><w:t>零</w:t></w:r></w:p>"
        '<w:p><w:commentRangeStart w:id="8"/><w:r><w:t>一</w:t></w:r>'
        '<w:commentRangeEnd w:id="8"/></w:p>'
    )
    docx = _write_docx(
        tmp_path,
        document_xml=document_xml,
        comments_xml=_comments_xml(("8", "DocWen", "批注")),
    )

    assert _extract_comments(docx) == [
        CommentAnchorInfo(
            comment_id="8",
            author="DocWen",
            text="批注",
            anchor_paragraph_index=1,
        )
    ]
