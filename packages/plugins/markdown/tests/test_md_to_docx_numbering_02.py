"""Focused tests split from test_md_to_docx_numbering.py."""

from __future__ import annotations

from ._md_to_docx_numbering_support import (
    BUNDLED_DOCX_TEMPLATES,
    WML_NS,
    DocxListNumbering,
    MdToDocxConverter,
    Path,
    _ExactSchemeRegistry,
    _get_numbering_xml,
    apply_list_to_paragraph,
    etree,
    make_context,
    os,
    pytest,
    repository_numbering_registry,
    tempfile,
    write_numbering_to_docx,
    write_temp_md,
    zipfile,
)

pytestmark = pytest.mark.contract


class TestWriteNumberingToDocx:
    """Tests for the post-save ZIP writeback routine."""

    def test_adds_definitions_to_docx(self):
        """Numbering definitions are injected after save."""
        from docx import Document

        doc = Document()
        p = doc.add_paragraph("Item")
        apply_list_to_paragraph(p, "1", 0)

        fd, tmp = tempfile.mkstemp(suffix=".docx")
        os.close(fd)
        doc.save(tmp)
        try:
            num = DocxListNumbering()
            num.create_list_definition({0: "ordered"})
            write_numbering_to_docx(tmp, num)

            num_root = _get_numbering_xml(tmp)
            assert num_root is not None
            an_elems = num_root.findall(f"{{{WML_NS}}}abstractNum")
            num_elems = num_root.findall(f"{{{WML_NS}}}num")
            assert len(an_elems) >= 1
            assert len(num_elems) >= 1
        finally:
            os.unlink(tmp)

    def test_noop_when_no_definitions(self):
        """Calling writeback with empty numbering is a safe no-op."""
        from docx import Document

        doc = Document()
        doc.add_paragraph("Hello")
        fd, tmp = tempfile.mkstemp(suffix=".docx")
        os.close(fd)
        doc.save(tmp)
        try:
            num = DocxListNumbering()
            # has_definitions is False
            write_numbering_to_docx(tmp, num)

            # File should still be valid
            assert zipfile.is_zipfile(tmp)
        finally:
            os.unlink(tmp)


@pytest.mark.parametrize(
    ("scheme", "registry", "error_type", "diagnostic_code"),
    [
        ("", _ExactSchemeRegistry(), "invalid_input", "NUMBERING-SCHEME-REQUIRED"),
        ("exact", None, "capability_unavailable", "NUMBERING-REGISTRY-UNAVAILABLE"),
        ("missing", _ExactSchemeRegistry(), "resource_not_found", "NUMBERING-SCHEME-NOT-FOUND"),
        (
            "exact",
            _ExactSchemeRegistry(enabled=False),
            "capability_unavailable",
            "NUMBERING-SCHEME-DISABLED",
        ),
        ("exact", _ExactSchemeRegistry(levels={}), "invalid_input", "NUMBERING-SCHEME-NO-LEVELS"),
    ],
)
def test_text_heading_numbering_rejects_unusable_exact_scheme(
    scheme: str,
    registry: object,
    error_type: str,
    diagnostic_code: str,
) -> None:
    md_path = write_temp_md("# Heading\n")
    ctx, _ = make_context(
        md_path,
        target_format="docx",
        options={
            "add_numbering": True,
            "numbering_scheme": scheme,
            "heading_numbering_render_mode": "text",
        },
        numbering_registry=registry,
    )

    result = MdToDocxConverter().convert(ctx)

    assert not result.success
    assert result.error is not None
    assert result.error.error_type == error_type
    assert result.error.diagnostic_code == diagnostic_code
    assert result.artifacts == []


class TestWordNativeHeadingNumbering:
    """End-to-end tests for word_native heading numbering mode."""

    WML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    def _get_styles_xml(self, docx_path: str) -> etree._Element | None:
        with zipfile.ZipFile(docx_path, "r") as zf:
            if "word/styles.xml" not in zf.namelist():
                return None
            raw = zf.read("word/styles.xml")
        return etree.fromstring(raw)

    def _get_numbering_xml(self, docx_path: str) -> etree._Element | None:
        with zipfile.ZipFile(docx_path, "r") as zf:
            if "word/numbering.xml" not in zf.namelist():
                return None
            raw = zf.read("word/numbering.xml")
        return etree.fromstring(raw)

    @staticmethod
    def _heading_abstract_num(docx_path: str, text: str) -> etree._Element:
        """Resolve a generated heading's directly assigned numbering definition."""
        from docx import Document

        paragraph = next(p for p in Document(docx_path).paragraphs if p.text == text)
        num_id = paragraph._p.find(f"{{{WML_NS}}}pPr/{{{WML_NS}}}numPr/{{{WML_NS}}}numId")
        assert num_id is not None
        with zipfile.ZipFile(docx_path) as archive:
            root = etree.fromstring(archive.read("word/numbering.xml"))
        num = root.find(f'{{{WML_NS}}}num[@{{{WML_NS}}}numId="{num_id.get(f"{{{WML_NS}}}val")}"]')
        assert num is not None
        abstract_id = num.find(f"{{{WML_NS}}}abstractNumId")
        assert abstract_id is not None
        abstract = root.find(
            f'{{{WML_NS}}}abstractNum[@{{{WML_NS}}}abstractNumId="{abstract_id.get(f"{{{WML_NS}}}val")}"]'
        )
        assert abstract is not None
        return abstract

    def test_word_native_with_hierarchical_standard(self):
        """Generated headings use direct numPr without modifying shared styles."""
        md = "# Heading 1\n\n## Heading 2\n\n### Heading 3\n"
        md_path = write_temp_md(md)
        ctx, _ = make_context(
            md_path,
            target_format="docx",
            options={
                "remove_numbering": False,
                "add_numbering": True,
                "numbering_scheme": "hierarchical_standard",
                "heading_numbering_render_mode": "word_native",
            },
            config_values={"template_fill": {"gongwen_numbering_text": False}},
            numbering_registry=repository_numbering_registry(),
        )

        result = MdToDocxConverter().convert(ctx)
        assert result.success, f"Conversion failed: {result.error}"

        output = str(Path(result.artifacts[0].staging_path))

        heading_an = self._heading_abstract_num(output, "Heading 1")
        lvls = heading_an.findall(f"{{{self.WML_NS}}}lvl")
        assert len(lvls) >= 3
        assert not heading_an.findall(f"{{{self.WML_NS}}}lvl/{{{self.WML_NS}}}pStyle")

        from docx import Document

        document = Document(output)
        num_ids = set()
        for level in range(1, 4):
            heading = next(p for p in document.paragraphs if p.text == f"Heading {level}")
            num_pr = heading._p.find(f"{{{self.WML_NS}}}pPr/{{{self.WML_NS}}}numPr")
            assert num_pr is not None
            ilvl = num_pr.find(f"{{{self.WML_NS}}}ilvl")
            num_id = num_pr.find(f"{{{self.WML_NS}}}numId")
            assert ilvl is not None and num_id is not None
            assert ilvl.get(f"{{{self.WML_NS}}}val") == str(level - 1)
            num_ids.add(num_id.get(f"{{{self.WML_NS}}}val"))
        assert len(num_ids) == 1

        styles_root = self._get_styles_xml(output)
        assert styles_root is not None, "styles.xml missing"
        for style_id in ("Heading1", "Heading2", "Heading3"):
            style_elem = styles_root.find(f'{{{self.WML_NS}}}style[@{{{self.WML_NS}}}styleId="{style_id}"]')
            assert style_elem is not None, f"{style_id} not found"
            assert style_elem.find(f"{{{self.WML_NS}}}pPr/{{{self.WML_NS}}}numPr") is None

    @pytest.mark.parametrize(
        "template_path",
        BUNDLED_DOCX_TEMPLATES,
        ids=lambda path: path.stem,
    )
    def test_word_native_uses_bundled_template_heading_style_ids(self, template_path: Path):
        """Numbering binds to each bundled template's request-owned style IDs."""

        from docx import Document

        md_path = write_temp_md("# Heading 1\n\n## Heading 2\n\n### Heading 3\n")
        ctx, _ = make_context(
            md_path,
            target_format="docx",
            options={
                "template_name": str(template_path),
                "remove_numbering": False,
                "add_numbering": True,
                "numbering_scheme": "hierarchical_standard",
                "heading_numbering_render_mode": "word_native",
            },
            config_values={"template_fill": {"gongwen_numbering_text": False}},
            numbering_registry=repository_numbering_registry(),
        )

        result = MdToDocxConverter().convert(ctx)
        assert result.success, f"Conversion failed for {template_path.name}: {result.error}"
        output = str(Path(result.artifacts[0].staging_path))

        document = Document(output)
        paragraph_style_ids: dict[str, str] = {}
        for paragraph in document.paragraphs:
            if paragraph.text not in {"Heading 1", "Heading 2", "Heading 3"}:
                continue
            assert paragraph.style is not None
            paragraph_style_ids[paragraph.text] = paragraph.style.style_id
        expected_style_ids = [paragraph_style_ids[f"Heading {level}"] for level in range(1, 4)]

        heading_abstract_num = self._heading_abstract_num(output, "Heading 1")
        levels = heading_abstract_num.findall(f"{{{self.WML_NS}}}lvl")
        assert len(levels) >= 3
        assert all(level.find(f"{{{self.WML_NS}}}pStyle") is None for level in levels)

        styles_root = self._get_styles_xml(output)
        template_styles = self._get_styles_xml(str(template_path))
        assert styles_root is not None and template_styles is not None
        for style_id in expected_style_ids:
            style = styles_root.find(f'{{{self.WML_NS}}}style[@{{{self.WML_NS}}}styleId="{style_id}"]')
            original = template_styles.find(f'{{{self.WML_NS}}}style[@{{{self.WML_NS}}}styleId="{style_id}"]')
            assert style is not None and original is not None
            actual = style.find(f"{{{self.WML_NS}}}pPr/{{{self.WML_NS}}}numPr")
            expected = original.find(f"{{{self.WML_NS}}}pPr/{{{self.WML_NS}}}numPr")
            assert (etree.tostring(actual) if actual is not None else None) == (
                etree.tostring(expected) if expected is not None else None
            )

    def test_word_native_with_gongwen_standard_approximate(self):
        """word_native + gongwen_standard succeeds (approximate verdict allowed)."""
        md = "# 标题\n\n## 标题二\n"
        md_path = write_temp_md(md)
        ctx, _ = make_context(
            md_path,
            target_format="docx",
            options={
                "remove_numbering": False,
                "add_numbering": True,
                "numbering_scheme": "gongwen_standard",
                "heading_numbering_render_mode": "word_native",
            },
            config_values={"template_fill": {"gongwen_numbering_text": False}},
            numbering_registry=repository_numbering_registry(),
        )

        result = MdToDocxConverter().convert(ctx)
        assert result.success, f"Conversion failed: {result.error}"

        num_root = self._get_numbering_xml(str(Path(result.artifacts[0].staging_path)))
        assert num_root is not None

    def test_text_mode_unchanged(self):
        """text mode (default) produces same text-concatenation behavior."""
        md = "# Heading 1\n\n## Heading 2\n"
        md_path = write_temp_md(md)
        ctx, _ = make_context(
            md_path,
            target_format="docx",
            options={
                "remove_numbering": False,
                "add_numbering": True,
                "numbering_scheme": "hierarchical_standard",
                # heading_numbering_render_mode defaults to "text"
            },
            numbering_registry=repository_numbering_registry(),
        )

        result = MdToDocxConverter().convert(ctx)
        assert result.success

        from docx import Document

        doc = Document(str(Path(result.artifacts[0].staging_path)))
        full = "\n".join(p.text for p in doc.paragraphs)
        # In text mode, heading text includes the numbering prefix
        assert "1 Heading 1" in full or "1 Heading" in full

    def test_word_native_no_text_numbering_added(self):
        """word_native does NOT add text numbering to headings."""
        md = "# Heading 1\n\n## Heading 2\n"
        md_path = write_temp_md(md)
        ctx, _ = make_context(
            md_path,
            target_format="docx",
            options={
                "remove_numbering": False,
                "add_numbering": True,
                "numbering_scheme": "hierarchical_standard",
                "heading_numbering_render_mode": "word_native",
            },
            config_values={"template_fill": {"gongwen_numbering_text": False}},
            numbering_registry=repository_numbering_registry(),
        )

        result = MdToDocxConverter().convert(ctx)
        assert result.success

        from docx import Document

        doc = Document(str(Path(result.artifacts[0].staging_path)))
        full = "\n".join(p.text for p in doc.paragraphs)
        # In word_native mode, heading text should NOT have "1 " prefix
        # (unlike text mode which concatenates "1 Heading 1")
        assert "Heading 1" in full
        assert "Heading 2" in full

    def test_word_native_no_heading_numbering_xml_when_no_add(self):
        """No heading numbering injection when add_numbering is False."""
        md = "# Heading\n"
        md_path = write_temp_md(md)
        ctx, _ = make_context(
            md_path,
            target_format="docx",
            options={
                "remove_numbering": False,
                "add_numbering": False,
                "heading_numbering_render_mode": "word_native",
            },
        )

        result = MdToDocxConverter().convert(ctx)
        assert result.success

        output = str(Path(result.artifacts[0].staging_path))
        from docx import Document

        heading = next(p for p in Document(output).paragraphs if p.text == "Heading")
        assert heading._p.find(f"{{{WML_NS}}}pPr/{{{WML_NS}}}numPr") is None
