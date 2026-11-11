"""Both DOCX export routes keep and authenticate template-owned caption identities."""

from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

from docwen_core._docx_semantics_v3_styles import parse_caption_style_binding_map
from docwen_core.docx_resolved_numbering_recovery import ResolvedNumberingV4Recovery
from docwen_core.docx_semantics_v3 import CAPTION_STYLE_BINDING_MAP_NAMESPACE, DocxSemanticsV3Recovery
from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter
from docwen_runtime.config.document_styles import build_document_style_catalog

from ._resolved_v4_support import _NEUTRAL, _PLAN, _context, _refs
from .conftest import PROJECT_ROOT, make_context

pytestmark = pytest.mark.contract


@pytest.mark.parametrize("route", ["v3", "v4"])
@pytest.mark.parametrize("style_id", ["9402", "图表_题注", '表&"<甲>'])
def test_caption_template_identity_survives_saved_conversion(tmp_path: Path, route: str, style_id: str) -> None:
    catalog = build_document_style_catalog(
        {}, request_options={"locale": "de_DE"}, locales_dir=PROJECT_ROOT / "i18n" / "locales"
    )
    document = Document()
    caption = document.styles.add_style(catalog.name_for("table_caption"), WD_STYLE_TYPE.PARAGRAPH)
    element = caption._element
    assert element is not None
    element.set(qn("w:styleId"), style_id)
    aliases = OxmlElement("w:aliases")
    aliases.set(qn("w:val"), "Table caption alias,表格别名")
    element.insert(1, aliases)
    original_style = etree.tostring(element, method="c14n", exclusive=True)
    document.add_paragraph("{{正文}}")
    template = tmp_path / "caption-template.docx"
    document.save(str(template))
    template_bytes = template.read_bytes()
    if route == "v3":
        source = tmp_path / "source.md"
        source.write_text("Table: Results ^results\n\n| A |\n|---|\n| 1 |\n", encoding="utf-8")
        context, _workspace = make_context(
            str(source),
            options={
                "template_name": str(template),
                "markdown_extensions": {"input": {"captions_references": True}},
            },
        )
    else:
        context, _workspace = _context(
            tmp_path,
            _refs(_NEUTRAL, _PLAN),
            options={"locale": "zh_CN", "heading_merge_mode": "never", "template_name": str(template)},
        )

    result = MdToDocxConverter().convert(context)

    assert result.success, (result.error, result.diagnostics)
    output = Path(result.artifacts[0].staging_path)
    with ZipFile(output) as package:
        styles = etree.fromstring(package.read("word/styles.xml"))
        preserved = next(item for item in styles.findall(qn("w:style")) if item.get(qn("w:styleId")) == style_id)
        assert etree.tostring(preserved, method="c14n", exclusive=True) == original_style
        payload = next(
            package.read(name)
            for name in package.namelist()
            if name.startswith("customXml/item")
            and name.endswith(".xml")
            and CAPTION_STYLE_BINDING_MAP_NAMESPACE.encode() in package.read(name)
        )
        binding = next(
            item
            for item in parse_caption_style_binding_map(etree.fromstring(payload))
            if item.semantic_key == "table_caption"
        )
        assert binding.resolved_style_id == style_id
        assert binding.visible_name == catalog.name_for("table_caption")
        content = etree.fromstring(package.read("word/document.xml"))
        assert any(item.get(qn("w:val")) == style_id for item in content.iter(qn("w:pStyle")))
    reopened = Document(str(output))
    if route == "v3":
        recovery = DocxSemanticsV3Recovery.load(output, reopened)
        assert recovery.caption_signatures == (("table", "results", "Results", "1"),)
    else:
        recovery = ResolvedNumberingV4Recovery.load_if_present(output, reopened)
        assert recovery is not None
        assert any(item[:3] == ("table", "results-main", "Results") for item in recovery.caption_signatures)
    assert template.read_bytes() == template_bytes
