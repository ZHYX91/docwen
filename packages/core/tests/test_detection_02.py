"""Focused tests split from test_detection.py."""

from __future__ import annotations

from ._detection_support import (
    StructureStatus,
    _ooxml_entries,
    _write_temp_file,
    _write_temp_text,
    _write_zip_entries,
    codecs,
    detect_content_format,
    detect_text_format,
    has_known_signature,
    is_text_file,
    os,
    pytest,
    tempfile,
)

pytestmark = pytest.mark.contract


class TestDetectTextFormat:
    """Text format detection by content analysis."""

    def test_detect_markdown_heading(self):
        """Markdown via heading patterns."""
        content = "# Title\n\nSome paragraph with **bold** text.\n"
        path = _write_temp_text(".txt", content)
        try:
            assert detect_text_format(path) == "markdown"
        finally:
            os.unlink(path)

    def test_detect_markdown_yaml_frontmatter(self):
        """Markdown via YAML front matter."""
        content = "---\ntitle: Test\n---\n\n# Heading\n\nContent.\n"
        path = _write_temp_text(".txt", content)
        try:
            assert detect_text_format(path) == "markdown"
        finally:
            os.unlink(path)

    def test_detect_markdown_code_fence(self):
        """Markdown via code fence."""
        content = "```python\nprint('hello')\n```\n"
        path = _write_temp_text(".txt", content)
        try:
            assert detect_text_format(path) == "markdown"
        finally:
            os.unlink(path)

    def test_detect_markdown_wiki_link(self):
        """Markdown via wiki link."""
        content = "See [[page|display]] for details.\n"
        path = _write_temp_text(".txt", content)
        try:
            assert detect_text_format(path) == "markdown"
        finally:
            os.unlink(path)

    def test_detect_markdown_table(self):
        """Markdown via table syntax."""
        content = "| A | B |\n| --- | --- |\n| 1 | 2 |\n"
        path = _write_temp_text(".txt", content)
        try:
            assert detect_text_format(path) == "markdown"
        finally:
            os.unlink(path)

    def test_detect_borderless_markdown_table_before_pipe_delimited_text(self):
        """A strict Markdown delimiter row outranks generic pipe-CSV sniffing."""
        content = "A | B\n--- | :---:\n1 | 2\n"
        path = _write_temp_text(".resource", content)
        try:
            assert detect_text_format(path) == "markdown"
            assert detect_content_format(path).format == "markdown"
        finally:
            os.unlink(path)

    @pytest.mark.parametrize(
        "content",
        [
            "A | B\n1 | 2\n",
            "A | B\n-- | ---\n1 | 2\n",
            "--- | ---\nA | B\n1 | 2\n",
        ],
    )
    def test_pipe_delimited_text_without_adjacent_strict_table_separator_stays_csv(self, content: str):
        path = _write_temp_text(".resource", content)
        try:
            assert detect_text_format(path) == "csv"
            assert detect_content_format(path).format == "csv"
        finally:
            os.unlink(path)

    def test_detect_csv(self):
        """CSV via comma delimiter analysis."""
        content = "a,b,c\n1,2,3\n4,5,6\n"
        path = _write_temp_text(".csv", content)
        try:
            assert detect_text_format(path) == "csv"
        finally:
            os.unlink(path)

    @pytest.mark.parametrize(
        "content",
        [
            "city;value\n北京;1\n上海;2\n",
            'city,notes\n北京,"line one\nline two"\n上海,done\n',
        ],
    )
    def test_detect_csv_uses_quote_aware_delimiter_sniffing(self, content: str):
        path = _write_temp_text(".csv", content)
        try:
            assert detect_text_format(path) == "csv"
            assert detect_content_format(path).format == "csv"
        finally:
            os.unlink(path)

    @pytest.mark.parametrize(
        "payload",
        [
            codecs.BOM_UTF16_LE + "city,value\n北京,1\n上海,2\n".encode("utf-16-le"),
            codecs.BOM_UTF16_BE + "city;value\n北京;1\n上海;2\n".encode("utf-16-be"),
        ],
    )
    def test_detect_utf16_csv_from_bom_content(self, payload: bytes):
        path = _write_temp_file(".csv", payload)
        try:
            assert is_text_file(path) is True
            detection = detect_content_format(path)
            assert detection.format == "csv"
            assert detection.method.value == "text_sniff"
        finally:
            os.unlink(path)

    def test_detect_tsv(self):
        """TSV via tab delimiter analysis."""
        content = "a\tb\tc\n1\t2\t3\n4\t5\t6\n"
        path = _write_temp_text(".tsv", content)
        try:
            assert detect_text_format(path) == "tsv"
        finally:
            os.unlink(path)

    def test_detect_html(self):
        """HTML via tag markers."""
        content = "<!DOCTYPE html>\n<html>\n<head><title>T</title></head>\n<body>B</body>\n</html>\n"
        path = _write_temp_text(".html", content)
        try:
            assert detect_text_format(path) == "html"
        finally:
            os.unlink(path)

    def test_detect_enex(self):
        """ENEX via en-export marker."""
        content = '<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE en-export>\n<en-export></en-export>\n'
        path = _write_temp_text(".enex", content)
        try:
            assert detect_text_format(path) == "enex"
        finally:
            os.unlink(path)

    def test_detect_mhtml(self):
        """MHTML via MIME headers."""
        content = (
            "MIME-Version: 1.0\r\n"
            'Content-Type: multipart/related; boundary="---boundary"\r\n'
            "\r\n"
            "<html><body>test</body></html>\r\n"
        )
        path = _write_temp_text(".mhtml", content)
        try:
            assert detect_text_format(path) == "mhtml"
        finally:
            os.unlink(path)

    def test_plain_text(self):
        """Plain text fallback."""
        content = "This is just a plain text file.\nNo special markers.\n"
        path = _write_temp_text(".txt", content)
        try:
            assert detect_text_format(path) == "txt"
        finally:
            os.unlink(path)


class TestIsTextFile:
    """Content-based text vs. binary classification."""

    def test_plain_text_is_text(self):
        path = _write_temp_text(".txt", "Hello, world!\nThis is text.\n" * 100)
        try:
            assert is_text_file(path) is True
        finally:
            os.unlink(path)

    def test_markdown_is_text(self):
        path = _write_temp_text(".md", "# Title\n\nParagraph with **bold**.\n")
        try:
            assert is_text_file(path) is True
        finally:
            os.unlink(path)

    def test_png_is_not_text(self):
        # Minimal valid-ish PNG header
        path = _write_temp_file(".png", b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR" + b"\x00" * 100)
        try:
            assert is_text_file(path) is False
        finally:
            os.unlink(path)

    def test_jpeg_is_not_text(self):
        path = _write_temp_file(".jpg", b"\xff\xd8\xff\xe0" + b"\x00" * 100)
        try:
            assert is_text_file(path) is False
        finally:
            os.unlink(path)

    def test_pdf_is_not_text(self):
        # PDFs with binary content
        path = _write_temp_file(".pdf", b"%PDF-1.4\n\x00\x01\x02\x03" + b"\x00" * 200)
        try:
            assert is_text_file(path) is False
        finally:
            os.unlink(path)

    def test_empty_file_is_text(self):
        path = _write_temp_file(".txt", b"")
        try:
            assert is_text_file(path) is True
        finally:
            os.unlink(path)

    def test_utf8_text_with_multibyte(self):
        path = _write_temp_text(".txt", "中文字符测试\n" * 50)
        try:
            assert is_text_file(path) is True
        finally:
            os.unlink(path)

    def test_gbk_text(self):
        content = "GBK编码测试文本\n" * 50
        path = _write_temp_file(".txt", content.encode("gbk"))
        try:
            assert is_text_file(path) is True
        finally:
            os.unlink(path)

    def test_null_bytes_are_binary(self):
        # File with >10% null bytes
        data = b"text" + b"\x00" * 100 + b"more"
        path = _write_temp_file(".bin", data)
        try:
            assert is_text_file(path) is False
        finally:
            os.unlink(path)


class TestHasKnownSignature:
    """Binary signature presence detection."""

    def test_pdf_has_signature(self):
        path = _write_temp_file(".pdf", b"%PDF-1.4\n")
        try:
            assert has_known_signature(path) is True
        finally:
            os.unlink(path)

    def test_png_has_signature(self):
        path = _write_temp_file(".png", b"\x89PNG\r\n\x1a\n")
        try:
            assert has_known_signature(path) is True
        finally:
            os.unlink(path)

    def test_markdown_no_signature(self):
        path = _write_temp_text(".md", "# Hello\n")
        try:
            assert has_known_signature(path) is False
        finally:
            os.unlink(path)

    def test_plain_text_no_signature(self):
        path = _write_temp_text(".txt", "plain text\n")
        try:
            assert has_known_signature(path) is False
        finally:
            os.unlink(path)

    def test_missing_file(self):
        assert has_known_signature("/nonexistent/file.xyz") is False


class TestFailClosedContentOutcomes:
    """Ambiguous content is never projected from the filename extension."""

    def test_unknown_binary_stays_unknown(self):
        path = _write_temp_file(".xyz", b"\x00\x01\x02\x03\x04\x05\x06\x07")
        try:
            result = detect_content_format(path)
            assert result.format == "unknown"
            assert result.detail_code == "FILE_CONTENT_UNRECOGNIZED"
        finally:
            os.unlink(path)

    def test_text_disguised_as_docx_is_still_text(self):
        path = _write_temp_file(".docx", b"Not a real docx file content")
        try:
            assert detect_content_format(path).format == "txt"
        finally:
            os.unlink(path)

    def test_empty_file_is_not_projected_as_text(self):
        path = _write_temp_file(".dat", b"")
        try:
            result = detect_content_format(path)
            assert result.format == "unknown"
            assert result.detail_code == "FILE_EMPTY"
            assert result.structure_status is StructureStatus.INVALID
        finally:
            os.unlink(path)


class TestErrorHandling:
    """Error and edge case handling."""

    def test_nonexistent_file_raises(self):
        with pytest.raises(FileNotFoundError):
            detect_content_format("/nonexistent/file_12345.xyz")


class TestCrossFormatDetection:
    """Detect format correctly regardless of file extension."""

    @pytest.mark.parametrize("brand", [b"heic", b"heix", b"hevc", b"hevx", b"mif1", b"msf1"])
    def test_heif_brand_with_misleading_suffix_is_detected_by_content(self, tmp_path, brand: bytes) -> None:
        source = tmp_path / f"{brand.decode('ascii')}.txt"
        source.write_bytes((24).to_bytes(4, "big") + b"ftyp" + brand + b"\x00\x00\x00\x00" + brand + b"\x00" * 4)

        assert detect_content_format(str(source)).format == "heic"

    def test_renamed_pdf_detected_by_content(self):
        """A PDF renamed to .txt is still detected as PDF."""
        path = _write_temp_file(".txt", b"%PDF-1.4\n%\x80\x80\x80\x80")
        try:
            # Content detection overrides extension
            assert detect_content_format(path).format == "pdf"
        finally:
            os.unlink(path)

    def test_renamed_png_detected_by_content(self):
        """A PNG renamed to .dat is still detected as PNG."""
        path = _write_temp_file(".dat", b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR")
        try:
            assert detect_content_format(path).format == "png"
        finally:
            os.unlink(path)

    def test_renamed_jpeg_detected_by_content(self):
        """A JPEG renamed to .txt is still detected as JPEG."""
        path = _write_temp_file(".txt", b"\xff\xd8\xff\xe0\x00\x10JFIF\x00")
        try:
            assert detect_content_format(path).format == "jpeg"
        finally:
            os.unlink(path)

    def test_renamed_docx_detected_by_zip_content(self):
        """A DOCX renamed to .zip is still detected as DOCX."""
        fd, path = tempfile.mkstemp(suffix=".zip")
        os.close(fd)
        try:
            _write_zip_entries(path, _ooxml_entries("docx"))
            assert detect_content_format(path).format == "docx"
        finally:
            os.unlink(path)
