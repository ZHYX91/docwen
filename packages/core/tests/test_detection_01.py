"""Focused tests split from test_detection.py."""

from __future__ import annotations

from ._detection_support import (
    SUPPORTED_EXTENSION_FORMATS,
    StructureStatus,
    _odf_entries,
    _ooxml_entries,
    _write_temp_file,
    _write_zip_entries,
    detect_content_format,
    os,
    pytest,
    tempfile,
    zipfile,
)

pytestmark = pytest.mark.contract


def test_supported_extension_registry_is_public_and_read_only() -> None:
    assert SUPPORTED_EXTENSION_FORMATS[".markdown"] == "markdown"
    with pytest.raises(TypeError):
        SUPPORTED_EXTENSION_FORMATS[".docwen-test"] = "txt"  # type: ignore[index]


class TestDetectContentFormat:
    """Content-based detection via magic bytes and container inspection."""

    def test_detect_jpeg(self):
        """JPEG magic bytes ``\\xff\\xd8\\xff``."""
        path = _write_temp_file(".jpg", b"\xff\xd8\xff\xe0\x00\x10JFIF\x00")
        try:
            assert detect_content_format(path).format == "jpeg"
        finally:
            os.unlink(path)

    def test_detect_png(self):
        """PNG magic bytes ``\\x89PNG\\r\\n\\x1a\\n``."""
        path = _write_temp_file(".png", b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR")
        try:
            assert detect_content_format(path).format == "png"
        finally:
            os.unlink(path)

    def test_detect_gif87a(self):
        """GIF87a magic bytes."""
        path = _write_temp_file(".gif", b"GIF87a\x00\x00\x00")
        try:
            assert detect_content_format(path).format == "gif"
        finally:
            os.unlink(path)

    def test_detect_gif89a(self):
        """GIF89a magic bytes."""
        path = _write_temp_file(".gif", b"GIF89a\x00\x00\x00")
        try:
            assert detect_content_format(path).format == "gif"
        finally:
            os.unlink(path)

    def test_detect_bmp(self):
        """BMP magic bytes ``BM``."""
        path = _write_temp_file(".bmp", b"BM\x00\x00\x00\x00\x00\x00\x00\x00")
        try:
            assert detect_content_format(path).format == "bmp"
        finally:
            os.unlink(path)

    def test_detect_tiff_le(self):
        """TIFF little-endian ``II*\\x00``."""
        path = _write_temp_file(".tiff", b"II*\x00\x08\x00\x00\x00")
        try:
            assert detect_content_format(path).format == "tiff"
        finally:
            os.unlink(path)

    def test_detect_tiff_be(self):
        """TIFF big-endian ``MM\\x00*``."""
        path = _write_temp_file(".tiff", b"MM\x00*\x00\x00\x00\x08")
        try:
            assert detect_content_format(path).format == "tiff"
        finally:
            os.unlink(path)

    def test_detect_pdf(self):
        """PDF magic bytes ``%PDF``."""
        path = _write_temp_file(".pdf", b"%PDF-1.4\n%\x80\x80\x80\x80")
        try:
            assert detect_content_format(path).format == "pdf"
        finally:
            os.unlink(path)

    def test_detect_rtf(self):
        """RTF magic bytes ``{\\rtf``."""
        path = _write_temp_file(".rtf", b"{\\rtf1\\ansi\\deff0")
        try:
            assert detect_content_format(path).format == "rtf"
        finally:
            os.unlink(path)

    def test_detect_webp(self):
        """WebP FourCC at offset 8 in RIFF container."""
        # RIFF header + "WEBP" at offset 8
        riff_header = b"RIFF\x00\x00\x00\x00WEBP"
        path = _write_temp_file(".webp", riff_header + b"\x00" * 20)
        try:
            assert detect_content_format(path).format == "webp"
        finally:
            os.unlink(path)


class TestDetectZipContainer:
    """ZIP-based container detection via member entry inspection."""

    def test_detect_docx(self):
        """A minimal valid OOXML Word package is detected as DOCX."""
        fd, path = tempfile.mkstemp(suffix=".docx")
        os.close(fd)
        try:
            _write_zip_entries(path, _ooxml_entries("docx"))
            assert detect_content_format(path).format == "docx"
        finally:
            os.unlink(path)

    def test_detect_xlsx(self):
        """A minimal valid OOXML spreadsheet package is detected as XLSX."""
        fd, path = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        try:
            _write_zip_entries(path, _ooxml_entries("xlsx"))
            assert detect_content_format(path).format == "xlsx"
        finally:
            os.unlink(path)

    def test_detect_pptx(self):
        """A minimal valid OOXML presentation package is detected as PPTX."""
        fd, path = tempfile.mkstemp(suffix=".pptx")
        os.close(fd)
        try:
            _write_zip_entries(path, _ooxml_entries("pptx"))
            assert detect_content_format(path).format == "pptx"
        finally:
            os.unlink(path)

    def test_detect_ofd(self):
        """OFD.xml and its safe DocRoot reference identify an OFD package."""
        fd, path = tempfile.mkstemp(suffix=".ofd")
        os.close(fd)
        try:
            _write_zip_entries(
                path,
                {
                    "OFD.xml": (
                        '<ofd:OFD xmlns:ofd="http://www.ofdspec.org/2016">'
                        "<ofd:DocBody><ofd:DocRoot>Doc_0/Document.xml</ofd:DocRoot></ofd:DocBody>"
                        "</ofd:OFD>"
                    ),
                    "Doc_0/Document.xml": '<ofd:Document xmlns:ofd="http://www.ofdspec.org/2016"/>',
                },
            )
            assert detect_content_format(path).format == "ofd"
        finally:
            os.unlink(path)

    def test_detect_epub(self):
        """An EPUB mimetype and referenced OPF package identify EPUB."""
        fd, path = tempfile.mkstemp(suffix=".epub")
        os.close(fd)
        try:
            _write_zip_entries(
                path,
                {
                    "mimetype": "application/epub+zip",
                    "META-INF/container.xml": (
                        '<container xmlns="urn:oasis:names:tc:opendocument:xmlns:container" version="1.0">'
                        '<rootfiles><rootfile full-path="EPUB/package.opf" '
                        'media-type="application/oebps-package+xml"/></rootfiles>'
                        "</container>"
                    ),
                    "EPUB/package.opf": '<package xmlns="http://www.idpf.org/2007/opf" version="3.0"/>',
                },
            )
            assert detect_content_format(path).format == "epub"
        finally:
            os.unlink(path)

    def test_detect_xps(self):
        """Content Types and a parseable fixed sequence identify XPS."""
        fd, path = tempfile.mkstemp(suffix=".xps")
        os.close(fd)
        try:
            _write_zip_entries(
                path,
                {
                    "[Content_Types].xml": (
                        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                        '<Default Extension="fdseq" '
                        'ContentType="application/vnd.ms-package.xps-fixeddocumentsequence+xml"/>'
                        "</Types>"
                    ),
                    "FixedDocumentSequence.fdseq": (
                        '<FixedDocumentSequence xmlns="http://schemas.microsoft.com/xps/2005/06"/>'
                    ),
                },
            )
            assert detect_content_format(path).format == "xps"
        finally:
            os.unlink(path)

    def test_detect_odt_via_mimetype(self):
        """A complete minimum ODF text package is detected as ODT."""
        fd, path = tempfile.mkstemp(suffix=".odt")
        os.close(fd)
        try:
            _write_zip_entries(path, _odf_entries("odt"))
            assert detect_content_format(path).format == "odt"
        finally:
            os.unlink(path)

    def test_detect_ods_via_mimetype(self):
        """A complete minimum ODF spreadsheet package is detected as ODS."""
        fd, path = tempfile.mkstemp(suffix=".ods")
        os.close(fd)
        try:
            _write_zip_entries(path, _odf_entries("ods"))
            assert detect_content_format(path).format == "ods"
        finally:
            os.unlink(path)

    def test_detect_odp_via_mimetype(self):
        """A complete minimum ODF presentation package is detected as ODP."""
        fd, path = tempfile.mkstemp(suffix=".odp")
        os.close(fd)
        try:
            _write_zip_entries(path, _odf_entries("odp"))
            assert detect_content_format(path).format == "odp"
        finally:
            os.unlink(path)

    @pytest.mark.parametrize(
        ("suffix", "marker"),
        [
            (".docx", "word/document.xml"),
            (".xlsx", "xl/workbook.xml"),
            (".pptx", "ppt/presentation.xml"),
            (".ofd", "OFD.xml"),
            (".epub", "META-INF/container.xml"),
            (".xps", "FixedDocumentSequence.fdseq"),
        ],
    )
    def test_single_marker_is_not_a_valid_document_package(self, suffix: str, marker: str):
        fd, path = tempfile.mkstemp(suffix=suffix)
        os.close(fd)
        try:
            _write_zip_entries(path, {marker: "<marker/>"})
            detection = detect_content_format(path)
            assert detection.structure_status is StructureStatus.INVALID
            assert detection.detail_code == "FILE_CONTAINER_INVALID"
        finally:
            os.unlink(path)

    def test_ooxml_requires_package_relationship(self):
        fd, path = tempfile.mkstemp(suffix=".docx")
        os.close(fd)
        try:
            entries = _ooxml_entries("docx")
            del entries["_rels/.rels"]
            _write_zip_entries(path, entries)
            detection = detect_content_format(path)
            assert detection.structure_status is StructureStatus.INVALID
            assert "_rels/.rels" in detection.detail_message
        finally:
            os.unlink(path)

    def test_ooxml_rejects_external_main_part_relationship(self):
        fd, path = tempfile.mkstemp(suffix=".docx")
        os.close(fd)
        try:
            entries = _ooxml_entries("docx")
            entries["_rels/.rels"] = (
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
                'TargetMode="External" Target="https://example.invalid/document.xml"/>'
                "</Relationships>"
            )
            _write_zip_entries(path, entries)
            detection = detect_content_format(path)
            assert detection.structure_status is StructureStatus.INVALID
            assert "must not be external" in detection.detail_message
        finally:
            os.unlink(path)

    def test_ooxml_rejects_malformed_required_xml(self):
        fd, path = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        try:
            entries = _ooxml_entries("xlsx")
            entries["xl/workbook.xml"] = "<workbook>"
            _write_zip_entries(path, entries)
            detection = detect_content_format(path)
            assert detection.structure_status is StructureStatus.INVALID
            assert "well-formed XML" in detection.detail_message
        finally:
            os.unlink(path)

    @pytest.mark.parametrize("missing_part", ["META-INF/manifest.xml", "content.xml"])
    def test_odf_requires_manifest_and_content(self, missing_part: str):
        fd, path = tempfile.mkstemp(suffix=".odt")
        os.close(fd)
        try:
            entries = _odf_entries("odt")
            del entries[missing_part]
            _write_zip_entries(path, entries)
            detection = detect_content_format(path)
            assert detection.structure_status is StructureStatus.INVALID
            assert missing_part in detection.detail_message
        finally:
            os.unlink(path)

    def test_odf_without_mimetype_is_not_identified_as_odf(self):
        fd, path = tempfile.mkstemp(suffix=".odt")
        os.close(fd)
        try:
            entries = _odf_entries("odt")
            del entries["mimetype"]
            _write_zip_entries(path, entries)
            detection = detect_content_format(path)
            assert detection.format == "zip"
            assert detection.detail_code == "FILE_CONTAINER_UNSUPPORTED"
        finally:
            os.unlink(path)

    def test_epub_requires_correct_mimetype_and_referenced_rootfile(self):
        fd, path = tempfile.mkstemp(suffix=".epub")
        os.close(fd)
        try:
            _write_zip_entries(
                path,
                {
                    "mimetype": "application/zip",
                    "META-INF/container.xml": (
                        '<container><rootfiles><rootfile full-path="missing.opf" '
                        'media-type="application/oebps-package+xml"/></rootfiles></container>'
                    ),
                },
            )
            detection = detect_content_format(path)
            assert detection.structure_status is StructureStatus.INVALID
            assert "application/epub+zip" in detection.detail_message
        finally:
            os.unlink(path)

    def test_epub_rejects_rootfile_reference_outside_the_package(self):
        fd, path = tempfile.mkstemp(suffix=".epub")
        os.close(fd)
        try:
            _write_zip_entries(
                path,
                {
                    "mimetype": "application/epub+zip",
                    "META-INF/container.xml": (
                        '<container><rootfiles><rootfile full-path="../outside.opf" '
                        'media-type="application/oebps-package+xml"/></rootfiles></container>'
                    ),
                },
            )
            detection = detect_content_format(path)
            assert detection.structure_status is StructureStatus.INVALID
            assert "escapes the archive" in detection.detail_message
        finally:
            os.unlink(path)

    def test_epub_requires_the_referenced_rootfile_to_exist(self):
        fd, path = tempfile.mkstemp(suffix=".epub")
        os.close(fd)
        try:
            _write_zip_entries(
                path,
                {
                    "mimetype": "application/epub+zip",
                    "META-INF/container.xml": (
                        '<container><rootfiles><rootfile full-path="EPUB/missing.opf" '
                        'media-type="application/oebps-package+xml"/></rootfiles></container>'
                    ),
                },
            )
            detection = detect_content_format(path)
            assert detection.structure_status is StructureStatus.INVALID
            assert "EPUB/missing.opf" in detection.detail_message
        finally:
            os.unlink(path)

    def test_ofd_requires_a_safe_referenced_docroot(self):
        fd, path = tempfile.mkstemp(suffix=".ofd")
        os.close(fd)
        try:
            _write_zip_entries(
                path,
                {
                    "OFD.xml": (
                        '<ofd:OFD xmlns:ofd="http://www.ofdspec.org/2016">'
                        "<ofd:DocBody><ofd:DocRoot>../Document.xml</ofd:DocRoot></ofd:DocBody>"
                        "</ofd:OFD>"
                    )
                },
            )
            detection = detect_content_format(path)
            assert detection.structure_status is StructureStatus.INVALID
            assert "escapes the archive" in detection.detail_message
        finally:
            os.unlink(path)

    def test_xps_requires_valid_content_types_declaration(self):
        fd, path = tempfile.mkstemp(suffix=".xps")
        os.close(fd)
        try:
            _write_zip_entries(
                path,
                {
                    "[Content_Types].xml": "<Types/>",
                    "FixedDocumentSequence.fdseq": "<FixedDocumentSequence/>",
                },
            )
            detection = detect_content_format(path)
            assert detection.structure_status is StructureStatus.INVALID
            assert "must declare" in detection.detail_message
        finally:
            os.unlink(path)

    def test_openxps_namespace_is_rejected_even_with_legacy_xps_content_type(self):
        fd, path = tempfile.mkstemp(suffix=".oxps")
        os.close(fd)
        try:
            _write_zip_entries(
                path,
                {
                    "[Content_Types].xml": (
                        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                        '<Default Extension="fdseq" '
                        'ContentType="application/vnd.ms-package.xps-fixeddocumentsequence+xml"/>'
                        "</Types>"
                    ),
                    "FixedDocumentSequence.fdseq": (
                        '<FixedDocumentSequence xmlns="http://schemas.openxps.org/oxps/v1.0"/>'
                    ),
                },
            )

            detection = detect_content_format(path)

            assert detection.structure_status is StructureStatus.INVALID
            assert detection.detail_code == "FILE_CONTAINER_INVALID"
            assert "OpenXPS is unsupported" in detection.detail_message
        finally:
            os.unlink(path)

    def test_unreadable_required_part_is_structurally_invalid(self):
        fd, path = tempfile.mkstemp(suffix=".docx")
        os.close(fd)
        try:
            _write_zip_entries(path, _ooxml_entries("docx"))
            with open(path, "rb") as stream:
                payload = bytearray(stream.read())
            marker = b"<document/>"
            offset = payload.index(marker)
            payload[offset + 1] ^= 1
            with open(path, "wb") as stream:
                stream.write(payload)

            detection = detect_content_format(path)
            assert detection.structure_status is StructureStatus.INVALID
            assert "cannot be read" in detection.detail_message
        finally:
            os.unlink(path)

    def test_unknown_zip_remains_an_explicit_unsupported_container(self):
        """A generic ZIP is never relabeled from its filename suffix."""
        fd, path = tempfile.mkstemp(suffix=".unknown_zip")
        os.close(fd)
        try:
            with zipfile.ZipFile(path, "w") as zf:
                zf.writestr("some_file.txt", "hello")
            result = detect_content_format(path)
            assert result.format == "zip"
            assert result.detail_code == "FILE_CONTAINER_UNSUPPORTED"
        finally:
            os.unlink(path)
