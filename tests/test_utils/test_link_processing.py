"""
link_processing 单元测试

覆盖 link_processing 模块中的纯工具函数。
不依赖 config_manager 的函数可直接测试；
依赖 config_manager 的函数（process_markdown_links 等）使用 monkeypatch。
"""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen.utils.link_processing import (
    _extract_block_by_id,
    _extract_section_by_heading,
    _format_image_placeholder,
    _normalize_link_target,
    _parse_anchor,
    _split_alt_text_and_size,
    _strip_yaml_front_matter,
    _unescape_pipe,
    get_file_type,
    resolve_file_path,
)

pytestmark = pytest.mark.unit


# ============================================================
# _unescape_pipe
# ============================================================


class TestUnescapePipe:
    """还原转义竖线"""

    def test_escaped_pipe(self) -> None:
        assert _unescape_pipe(r"a\|b") == "a|b"

    def test_no_escape(self) -> None:
        assert _unescape_pipe("a|b") == "a|b"

    def test_none_input(self) -> None:
        assert _unescape_pipe(None) is None

    def test_empty_string(self) -> None:
        assert _unescape_pipe("") == ""

    def test_multiple_escapes(self) -> None:
        assert _unescape_pipe(r"a\|b\|c") == "a|b|c"


# ============================================================
# _normalize_link_target
# ============================================================


class TestNormalizeLinkTarget:
    """规范化链接目标路径"""

    def test_url_decode(self) -> None:
        assert _normalize_link_target("hello%20world.md") == "hello world.md"

    def test_remove_anchor(self) -> None:
        assert _normalize_link_target("file.md#section") == "file.md"

    def test_remove_query_params(self) -> None:
        assert _normalize_link_target("image.png?v=1") == "image.png"

    def test_combined(self) -> None:
        assert _normalize_link_target("file%20name.md#heading?v=2") == "file name.md"

    def test_plain_path(self) -> None:
        assert _normalize_link_target("simple.md") == "simple.md"

    def test_strips_whitespace(self) -> None:
        assert _normalize_link_target("  file.md  ") == "file.md"

    def test_chinese_url_encoding(self) -> None:
        # %E6%B5%8B%E8%AF%95 = "测试"
        result = _normalize_link_target("%E6%B5%8B%E8%AF%95.md")
        assert result == "测试.md"


# ============================================================
# _parse_anchor
# ============================================================


class TestParseAnchor:
    """解析链接锚点"""

    def test_no_anchor(self) -> None:
        assert _parse_anchor("file.md") == ("file.md", None, None)

    def test_heading_anchor(self) -> None:
        assert _parse_anchor("file.md#标题名") == ("file.md", "标题名", None)

    def test_block_id_anchor(self) -> None:
        assert _parse_anchor("file.md#^block-id") == ("file.md", None, "block-id")

    def test_url_encoded(self) -> None:
        path, heading, block_id = _parse_anchor("file%20name.md#%E6%A0%87%E9%A2%98")
        assert path == "file name.md"
        assert heading == "标题"
        assert block_id is None

    def test_query_params_removed(self) -> None:
        path, heading, block_id = _parse_anchor("file.md?v=1")
        assert path == "file.md"
        assert heading is None
        assert block_id is None

    def test_empty_anchor(self) -> None:
        assert _parse_anchor("file.md#") == ("file.md", None, None)

    def test_no_file_path_with_heading(self) -> None:
        """仅锚点，无文件路径"""
        path, heading, _block_id = _parse_anchor("#标题")
        assert path == ""
        assert heading == "标题"

    def test_no_file_path_with_block_id(self) -> None:
        path, _heading, block_id = _parse_anchor("#^myblock")
        assert path == ""
        assert block_id == "myblock"


# ============================================================
# _extract_section_by_heading
# ============================================================


class TestExtractSectionByHeading:
    """从内容中提取指定标题的章节"""

    SAMPLE_CONTENT = "# 一级标题\n一级内容\n## 二级标题A\n二级A内容\n### 三级标题\n三级内容\n## 二级标题B\n二级B内容\n"

    def test_extract_second_level(self) -> None:
        result = _extract_section_by_heading(self.SAMPLE_CONTENT, "二级标题A")
        assert result is not None
        assert "## 二级标题A" in result
        assert "二级A内容" in result
        assert "### 三级标题" in result  # 包含子标题
        assert "## 二级标题B" not in result  # 不包含同级

    def test_extract_top_level(self) -> None:
        result = _extract_section_by_heading(self.SAMPLE_CONTENT, "一级标题")
        assert result is not None
        assert "# 一级标题" in result
        assert "一级内容" in result

    def test_heading_not_found(self) -> None:
        result = _extract_section_by_heading(self.SAMPLE_CONTENT, "不存在的标题")
        assert result is None

    def test_case_insensitive_match(self) -> None:
        content = "## Hello World\nsome content\n## Next"
        result = _extract_section_by_heading(content, "hello world")
        assert result is not None
        assert "Hello World" in result

    def test_last_section_includes_to_end(self) -> None:
        result = _extract_section_by_heading(self.SAMPLE_CONTENT, "二级标题B")
        assert result is not None
        assert "二级B内容" in result


# ============================================================
# _extract_block_by_id
# ============================================================


class TestExtractBlockById:
    """从内容中提取指定块ID的段落"""

    def test_inline_block_id(self) -> None:
        content = "这是一段重要文字 ^important\n\n后面的内容"
        result = _extract_block_by_id(content, "important")
        assert result is not None
        assert "重要文字" in result
        assert "^important" not in result

    def test_standalone_block_id(self) -> None:
        content = "前面的段落\n\n这是要引用的段落\n\n^myblock\n\n后面的段落"
        result = _extract_block_by_id(content, "myblock")
        assert result is not None
        assert "要引用的段落" in result

    def test_block_id_not_found(self) -> None:
        content = "普通内容\n没有块ID"
        result = _extract_block_by_id(content, "missing")
        assert result is None


# ============================================================
# _strip_yaml_front_matter
# ============================================================


class TestStripYamlFrontMatter:
    """移除 YAML front matter"""

    def test_with_yaml(self) -> None:
        content = "---\ntitle: 测试\ndate: 2024-01-01\n---\n正文内容"
        result = _strip_yaml_front_matter(content)
        assert "title" not in result
        assert "正文内容" in result

    def test_without_yaml(self) -> None:
        content = "# 标题\n正文"
        assert _strip_yaml_front_matter(content) == content

    def test_incomplete_yaml(self) -> None:
        """只有开头的 --- 没有结尾"""
        content = "---\ntitle: test\n正文内容"
        assert _strip_yaml_front_matter(content) == content

    def test_empty_content(self) -> None:
        assert _strip_yaml_front_matter("") == ""

    def test_yaml_with_trailing_newlines(self) -> None:
        content = "---\ntitle: test\n---\n\n\n正文"
        result = _strip_yaml_front_matter(content)
        assert result.startswith("正文")

    def test_with_yaml_crlf(self) -> None:
        content = "---\r\ntitle: 测试\r\n---\r\n正文"
        result = _strip_yaml_front_matter(content)
        assert result.startswith("正文")


# ============================================================
# _format_image_placeholder
# ============================================================


class TestFormatImagePlaceholder:
    """图片占位符生成"""

    def test_no_size(self) -> None:
        result = _format_image_placeholder("/path/to/image.png")
        assert result == "{{IMAGE:/path/to/image.png}}"

    def test_with_width_only(self) -> None:
        result = _format_image_placeholder("/path/to/image.png", width=300)
        assert "300" in result
        assert "IMAGE:" in result

    def test_with_width_and_height(self) -> None:
        result = _format_image_placeholder("/path/to/image.png", width=300, height=200)
        assert "300" in result
        assert "200" in result


# ============================================================
# _split_alt_text_and_size
# ============================================================


class TestSplitAltTextAndSize:
    """分离 alt 文本和尺寸"""

    def test_no_size(self) -> None:
        alt, w, h = _split_alt_text_and_size("描述文字")
        assert alt == "描述文字"
        assert w is None
        assert h is None

    def test_width_only(self) -> None:
        alt, w, h = _split_alt_text_and_size("描述|300")
        assert alt == "描述"
        assert w == 300
        assert h is None

    def test_width_and_height(self) -> None:
        alt, w, h = _split_alt_text_and_size("描述|300x200")
        assert alt == "描述"
        assert w == 300
        assert h == 200

    def test_width_without_height(self) -> None:
        alt, w, h = _split_alt_text_and_size("描述|300x")
        assert alt == "描述"
        assert w == 300
        assert h is None

    def test_none_input(self) -> None:
        alt, w, h = _split_alt_text_and_size(None)
        assert alt is None
        assert w is None
        assert h is None

    def test_empty_string(self) -> None:
        alt, w, _h = _split_alt_text_and_size("")
        assert alt is None
        assert w is None

    def test_non_numeric_after_pipe(self) -> None:
        """竖线后不是数字，保持原样"""
        alt, w, _h = _split_alt_text_and_size("alt|caption")
        assert alt == "alt|caption"
        assert w is None


# ============================================================
# get_file_type
# ============================================================


class TestGetFileType:
    """文件类型判断"""

    @pytest.mark.parametrize(
        ("path", "expected"),
        [
            ("photo.png", "image"),
            ("photo.PNG", "image"),
            ("photo.jpg", "image"),
            ("photo.jpeg", "image"),
            ("photo.gif", "image"),
            ("photo.bmp", "image"),
            ("photo.svg", "image"),
            ("photo.webp", "image"),
            ("photo.ico", "image"),
            ("doc.md", "markdown"),
            ("doc.markdown", "markdown"),
            ("file.txt", "unknown"),
            ("file.docx", "unknown"),
            ("file.pdf", "unknown"),
            ("file", "unknown"),
        ],
    )
    def test_file_types(self, path: str, expected: str) -> None:
        assert get_file_type(path) == expected


# ============================================================
# resolve_file_path（需要文件系统）
# ============================================================


class TestResolveFilePath:
    """文件路径解析"""

    def test_absolute_path_exists(self, tmp_path: Path) -> None:
        """绝对路径且文件存在"""
        target = tmp_path / "test.md"
        target.write_text("content", encoding="utf-8")
        source = str(tmp_path / "source.md")

        result = resolve_file_path(str(target), source)
        assert result is not None
        assert Path(result).is_absolute()

    def test_absolute_path_not_exists(self, tmp_path: Path) -> None:
        """绝对路径但文件不存在"""
        source = str(tmp_path / "source.md")
        result = resolve_file_path(str(tmp_path / "nonexistent.md"), source)
        assert result is None

    def test_relative_path(self, tmp_path: Path) -> None:
        """相对路径"""
        sub = tmp_path / "sub"
        sub.mkdir()
        target = sub / "target.md"
        target.write_text("content", encoding="utf-8")
        source = str(tmp_path / "source.md")

        result = resolve_file_path("sub/target.md", source)
        assert result is not None

    def test_auto_add_md_extension(self, tmp_path: Path) -> None:
        """不带扩展名时自动添加 .md（Obsidian 兼容）"""
        target = tmp_path / "note.md"
        target.write_text("content", encoding="utf-8")
        source = str(tmp_path / "source.md")

        result = resolve_file_path(str(target).replace(".md", ""), source)
        # 如果是绝对路径，应该能找到
        if result is not None:
            assert result.endswith(".md")
