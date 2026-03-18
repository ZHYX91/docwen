from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from docwen.docx_spell.api import build_anchor_report_markdown, redact_text

pytestmark = pytest.mark.unit


_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _write_min_docx(tmp_path: Path, *, document_xml: str, comments_xml: str | None) -> Path:
    out = tmp_path / "in.docx"
    with zipfile.ZipFile(out, "w") as zf:
        zf.writestr("word/document.xml", document_xml.encode("utf-8"))
        if comments_xml is not None:
            zf.writestr("word/comments.xml", comments_xml.encode("utf-8"))
    return out


def test_anchor_report_extracts_covered_text_and_comment_text(tmp_path: Path) -> None:
    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{_W_NS}">
  <w:body>
    <w:p>
      <w:r><w:t>前</w:t></w:r>
      <w:commentRangeStart w:id="0"/>
      <w:r><w:t>错</w:t></w:r>
      <w:commentRangeEnd w:id="0"/>
      <w:r><w:t>后</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
"""
    comments_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="{_W_NS}">
  <w:comment w:id="0">
    <w:p><w:r><w:t>测试批注</w:t></w:r></w:p>
  </w:comment>
</w:comments>
"""
    docx = _write_min_docx(tmp_path, document_xml=document_xml, comments_xml=comments_xml)
    md = build_anchor_report_markdown(docx, context_chars=10, redact=False)

    assert "批注 `0`" in md
    assert "批注内容：测试批注" in md
    assert "[错]" in md


def test_anchor_report_redacts_content(tmp_path: Path) -> None:
    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{_W_NS}">
  <w:body>
    <w:p>
      <w:r><w:t>ABC123中文</w:t></w:r>
      <w:commentRangeStart w:id="1"/>
      <w:r><w:t>错</w:t></w:r>
      <w:commentRangeEnd w:id="1"/>
    </w:p>
  </w:body>
</w:document>
"""
    comments_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="{_W_NS}">
  <w:comment w:id="1">
    <w:p><w:r><w:t>敏感内容ABC123</w:t></w:r></w:p>
  </w:comment>
</w:comments>
"""
    docx = _write_min_docx(tmp_path, document_xml=document_xml, comments_xml=comments_xml)
    md = build_anchor_report_markdown(docx, context_chars=10, redact=True)

    assert "敏感内容" not in md
    assert "ABC123" not in md
    assert "[█]" in md


def test_anchor_report_handles_missing_comments_xml(tmp_path: Path) -> None:
    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{_W_NS}">
  <w:body>
    <w:p>
      <w:commentRangeStart w:id="2"/>
      <w:r><w:t>错</w:t></w:r>
      <w:commentRangeEnd w:id="2"/>
    </w:p>
  </w:body>
</w:document>
"""
    docx = _write_min_docx(tmp_path, document_xml=document_xml, comments_xml=None)
    md = build_anchor_report_markdown(docx, context_chars=10, redact=False)

    assert "批注数（comments.xml）：`0`" in md
    assert "未找到批注文案的锚点 ID" in md
    assert "- `2`" in md


def test_anchor_report_multiple_comments_in_one_paragraph(tmp_path: Path) -> None:
    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{_W_NS}">
  <w:body>
    <w:p>
      <w:r><w:t>A</w:t></w:r>
      <w:commentRangeStart w:id="0"/>
      <w:r><w:t>X</w:t></w:r>
      <w:commentRangeEnd w:id="0"/>
      <w:r><w:t>B</w:t></w:r>
      <w:commentRangeStart w:id="1"/>
      <w:r><w:t>Y</w:t></w:r>
      <w:commentRangeEnd w:id="1"/>
      <w:r><w:t>C</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
"""
    comments_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="{_W_NS}">
  <w:comment w:id="0"><w:p><w:r><w:t>C0</w:t></w:r></w:p></w:comment>
  <w:comment w:id="1"><w:p><w:r><w:t>C1</w:t></w:r></w:p></w:comment>
</w:comments>
"""
    docx = _write_min_docx(tmp_path, document_xml=document_xml, comments_xml=comments_xml)
    md = build_anchor_report_markdown(docx, context_chars=10, redact=False)

    assert "批注 `0`" in md
    assert "批注 `1`" in md
    assert "[X]" in md
    assert "[Y]" in md


def test_anchor_report_empty_range_does_not_crash(tmp_path: Path) -> None:
    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{_W_NS}">
  <w:body>
    <w:p>
      <w:r><w:t>前</w:t></w:r>
      <w:commentRangeStart w:id="3"/>
      <w:commentRangeEnd w:id="3"/>
      <w:r><w:t>后</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
"""
    comments_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="{_W_NS}">
  <w:comment w:id="3"><w:p><w:r><w:t>空范围</w:t></w:r></w:p></w:comment>
</w:comments>
"""
    docx = _write_min_docx(tmp_path, document_xml=document_xml, comments_xml=comments_xml)
    md = build_anchor_report_markdown(docx, context_chars=10, redact=False)

    assert "批注 `3`" in md
    assert "[]" in md


def test_anchor_report_cross_paragraph_is_reported(tmp_path: Path) -> None:
    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{_W_NS}">
  <w:body>
    <w:p>
      <w:r><w:t>段一</w:t></w:r>
      <w:commentRangeStart w:id="4"/>
    </w:p>
    <w:p>
      <w:r><w:t>段二</w:t></w:r>
      <w:commentRangeEnd w:id="4"/>
    </w:p>
  </w:body>
</w:document>
"""
    comments_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="{_W_NS}">
  <w:comment w:id="4"><w:p><w:r><w:t>跨段</w:t></w:r></w:p></w:comment>
</w:comments>
"""
    docx = _write_min_docx(tmp_path, document_xml=document_xml, comments_xml=comments_xml)
    md = build_anchor_report_markdown(docx, context_chars=10, redact=False)

    assert "锚点异常（跨段/未闭合）" in md
    assert "跨段锚点" in md
    assert "`4`：段落 `1`" in md


@pytest.mark.parametrize(
    "text, expected",
    [
        # ASCII 字母数字
        ("Hello123", "████████"),
        # 中文（CJK 基本区）
        ("你好世界", "████"),
        # 日文平假名 + 片假名
        ("あいうアイウ", "██████"),
        # 韩文音节
        ("한국어", "███"),
        # 西里尔字母（俄文）
        ("Привет", "██████"),
        # 拉丁扩展（法/德/西/葡重音字母）
        ("àéîöüñ", "██████"),
        # 越南文拉丁扩展
        ("ếồắ", "███"),
        # 全角字母数字
        ("Ａ１ｂ２", "████"),
        # 标点符号应保留
        ("（）、。！？「」", "（）、。！？「」"),
        # 混合文本：字符被替换，标点保留
        ("张三（test）", "██（████）"),
    ],
    ids=[
        "ascii",
        "cjk_basic",
        "japanese_kana",
        "korean_syllables",
        "cyrillic",
        "latin_accented",
        "vietnamese",
        "fullwidth",
        "punctuation_preserved",
        "mixed",
    ],
)
def test_redact_text_covers_multilingual_characters(text: str, expected: str) -> None:
    assert redact_text(text) == expected
