"""Saved DOCX comments cover table paragraphs once, in document order."""

from __future__ import annotations

import zipfile
from xml.etree import ElementTree as ET

import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import parse_xml

from docwen_core.models.proofread import ProofreadRules
from docwen_plugin_proofread.docx_validator import DocxValidator, _iter_main_story_paragraphs

from ._proofread_plugin_support import _build_fake_context

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _validate(document, tmp_path, **options):
    source = tmp_path / "source.docx"
    document.save(source)
    original = source.read_bytes()
    staging = tmp_path / "staging"
    staging.mkdir()
    context = _build_fake_context(
        str(source),
        str(staging),
        target_format="docx",
        action_name="validate",
        options={
            "locale": "zh_CN",
            "enable_symbol_pairing": True,
            "enable_symbol_correction": False,
            "enable_typos_rule": False,
            "enable_sensitive_word": False,
            **options,
        },
        proofread_rules=ProofreadRules(symbol_pairs=(("（", "）"),)),
    )
    result = DocxValidator().convert(context)
    assert result.success, result.error
    assert source.read_bytes() == original
    artifact = result.artifacts[0]
    assert artifact.metadata["comments_failed"] == 0
    with zipfile.ZipFile(artifact.staging_path) as package:
        xml = ET.fromstring(package.read("word/document.xml"))
        comments = ET.fromstring(package.read("word/comments.xml")) if "word/comments.xml" in package.namelist() else []
    return result, xml, comments


def _anchored_paragraphs(xml):
    found = []
    for paragraph in xml.iter(W + "p"):
        starts = list(paragraph.iter(W + "commentRangeStart"))
        if not starts:
            continue
        assert len(starts) == 1
        identifier = starts[0].get(W + "id")
        active = False
        covered = ""
        for node in paragraph.iter():
            if node.tag == W + "commentRangeStart":
                active = True
            elif node.tag == W + "commentRangeEnd":
                assert node.get(W + "id") == identifier
                active = False
            elif active and node.tag == W + "t":
                covered += node.text or ""
        assert covered == "（"
        assert not active
        assert [ref.get(W + "id") for ref in paragraph.iter(W + "commentReference")] == [identifier]
        found.append(("".join(paragraph.itertext()), identifier))
    return found


@pytest.mark.parametrize("rows,cols", [(1, 2), (2, 1), (2, 2)])
def test_body_plain_nested_and_merged_tables_have_exact_ordered_comments(tmp_path, rows, cols):
    document = Document()
    document.add_paragraph("正文（")
    document.add_table(rows=1, cols=1).cell(0, 0).text = "普通（"
    table = document.add_table(rows=rows, cols=cols)
    merged = table.cell(0, 0).merge(table.cell(rows - 1, cols - 1))
    merged.text = "合并（"
    merged.add_paragraph("表前（")
    merged.add_table(rows=1, cols=1).cell(0, 0).text = "嵌套（"
    merged.paragraphs[-1].text = "表后（"
    document.add_paragraph("尾部（")

    result, xml, comments = _validate(document, tmp_path)

    expected = ["正文（", "普通（", "合并（", "表前（", "嵌套（", "表后（", "尾部（"]
    assert _anchored_paragraphs(xml) == [(text, str(index)) for index, text in enumerate(expected)]
    assert len(comments) == 7
    assert {comment.get(W + "author") for comment in comments} == {"DocWen"}
    metadata = result.artifacts[0].metadata
    assert {
        key: metadata[key] for key in ("paragraphs_total", "paragraphs_checked", "errors_found", "comments_added")
    } == dict.fromkeys(("paragraphs_total", "paragraphs_checked", "errors_found", "comments_added"), 7)
    assert metadata["paragraphs_skipped"] == 0
    assert result.metrics.extra["paragraphs_total"] == 7
    assert "7 paragraph(s)" in result.diagnostics[0].message


@pytest.mark.parametrize("skip", [True, False])
def test_table_only_document_preserves_code_quote_skips_and_counts(tmp_path, skip):
    document = Document()
    document.styles.add_style("Code Block", WD_STYLE_TYPE.PARAGRAPH)
    cell = document.add_table(rows=1, cols=1).cell(0, 0)
    cell.paragraphs[0].text = "代码（"
    cell.paragraphs[0].style = "Code Block"
    cell.add_paragraph("引用（", style="Quote")
    cell.add_table(rows=1, cols=1).cell(0, 0).text = "嵌套（"
    result, xml, comments = _validate(document, tmp_path, skip_code_blocks=skip, skip_quote_blocks=skip)
    expected = ["嵌套（"] if skip else ["代码（", "引用（", "嵌套（"]
    assert [text for text, _identifier in _anchored_paragraphs(xml)] == expected
    assert len(comments) == len(expected)
    metadata = result.artifacts[0].metadata
    assert metadata["paragraphs_total"] == 4  # Includes the required empty paragraph after the nested table.
    assert metadata["paragraphs_checked"] == len(expected)
    assert metadata["paragraphs_skipped"] == (2 if skip else 0)


def test_headers_footers_and_textboxes_remain_outside_body_scope(tmp_path):
    document = Document()
    document.add_paragraph("正文正常")
    document.sections[0].header.paragraphs[0].text = "页眉（"
    document.sections[0].footer.paragraphs[0].text = "页脚（"
    document.add_paragraph()._p.append(
        parse_xml(
            '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:v="urn:schemas-microsoft-com:vml"><w:pict><v:shape><w:txbxContent>'
            "<w:p><w:r><w:t>文本框（</w:t></w:r></w:p>"
            "</w:txbxContent></v:shape></w:pict></w:r>"
        )
    )
    result, _xml, comments = _validate(document, tmp_path)
    assert not comments
    metadata = result.artifacts[0].metadata
    assert metadata["paragraphs_total"] == 2
    assert metadata["paragraphs_checked"] == 1
    assert metadata["errors_found"] == 0


def test_table_walk_observes_cancellation_before_entering_nested_cells():
    document = Document()
    cell = document.add_table(rows=1, cols=1).cell(0, 0)
    cell.text = "首段"
    cell.add_table(rows=1, cols=1).cell(0, 0).text = "不能检查"
    cancelled = False

    def check():
        if cancelled:
            raise RuntimeError("cancelled")

    paragraphs = _iter_main_story_paragraphs(document, check)
    assert next(paragraphs).text == "首段"
    cancelled = True
    with pytest.raises(RuntimeError, match="cancelled"):
        next(paragraphs)
