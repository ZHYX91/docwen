"""Template style identities remain valid for Word's name-based field references."""

from __future__ import annotations

from copy import deepcopy
from io import BytesIO
from pathlib import Path
from typing import TYPE_CHECKING, cast
from zipfile import ZipFile

import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.styles.style import ParagraphStyle
from lxml import etree

from docwen_core.docx_styles import DocumentStyleCatalog
from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter
from docwen_plugin_markdown.to_docx.managed_styles import (
    ManagedStyleCompletionError,
    complete_managed_styles,
    validate_managed_style_package,
)
from docwen_runtime.config.document_styles import build_document_style_catalog

from .conftest import PROJECT_ROOT, make_context

if TYPE_CHECKING:
    from docx.document import Document as DocumentObject
    from docx.text.paragraph import Paragraph

pytestmark = pytest.mark.contract

_BODY_STYLE_ID = "9401"
_BODY_ALIAS = "Archive Body"
_FIELD_TAGS = frozenset(qn(f"w:{name}") for name in ("fldSimple", "fldChar", "instrText"))


def _catalog(locale: str) -> DocumentStyleCatalog:
    return build_document_style_catalog(
        {}, request_options={"locale": locale}, locales_dir=PROJECT_ROOT / "i18n" / "locales"
    )


def _xml(element: etree._Element) -> bytes:
    return etree.tostring(element, method="c14n", exclusive=True, with_tail=False)


def _style(root: etree._Element, style_id: str) -> etree._Element:
    return next(element for element in root.findall(qn("w:style")) if element.get(qn("w:styleId")) == style_id)


def _complex_field(paragraph: Paragraph, instruction_parts: tuple[str, ...]) -> None:
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    paragraph.add_run()._r.append(begin)
    for text in instruction_parts:
        instruction = OxmlElement("w:instrText")
        instruction.set(qn("xml:space"), "preserve")
        instruction.text = text
        paragraph.add_run()._r.append(instruction)
    separator = OxmlElement("w:fldChar")
    separator.set(qn("w:fldCharType"), "separate")
    paragraph.add_run()._r.append(separator)
    paragraph.add_run("Template field result")
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    paragraph.add_run()._r.append(end)


def _simple_field(paragraph: Paragraph, instruction: str) -> None:
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), instruction)
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "Template field result"
    run.append(text)
    field.append(run)
    paragraph._p.append(field)


def _template(locale: str) -> DocumentObject:
    document = Document()
    name = _catalog(locale).name_for("body_paragraph")
    body = cast(ParagraphStyle, document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH))
    body.font.size = Pt(10.5)
    body.paragraph_format.space_after = Pt(7)
    element = body._element
    assert element is not None
    element.set(qn("w:styleId"), _BODY_STYLE_ID)
    aliases = OxmlElement("w:aliases")
    aliases.set(qn("w:val"), f"{_BODY_ALIAS},正文别名")
    element.insert(1, aliases)
    if locale == "ar_SA":
        element.get_or_add_pPr().append(OxmlElement("w:bidi"))
        element.get_or_add_rPr().append(OxmlElement("w:rtl"))
    document.add_paragraph("Fixed template paragraph", style=body)
    _simple_field(document.add_paragraph(), f' TOC \\t "{name},1,{_BODY_ALIAS},2" \\h ')
    _complex_field(document.sections[0].header.paragraphs[0], (" STY", f'LEREF "{name}"', " \\* MERGEFORMAT "))
    _simple_field(document.sections[0].footer.paragraphs[0], f' STYLEREF "{_BODY_ALIAS}" \\* MERGEFORMAT ')
    document.add_paragraph("{{正文}}", style=body)
    return document


def _field_paragraphs(root: etree._Element) -> tuple[bytes, ...]:
    return tuple(
        _xml(paragraph)
        for paragraph in root.iter(qn("w:p"))
        if any(element.tag in _FIELD_TAGS for element in paragraph.iter())
    )


@pytest.mark.parametrize(("template_locale", "output_locale"), [("de_DE", "zh_CN"), ("ar_SA", "en_US")])
def test_converter_preserves_localized_numeric_style_and_dependent_fields(
    tmp_path: Path, template_locale: str, output_locale: str
) -> None:
    document = _template(template_locale)
    template = tmp_path / "field-template.docx"
    document.save(str(template))
    original = template.read_bytes()
    source = tmp_path / "source.md"
    source.write_text("New generated body.\n", encoding="utf-8")
    context, _workspace = make_context(
        str(source),
        options={"template_name": str(template)},
        document_style_catalog=_catalog(output_locale),
    )

    result = MdToDocxConverter().convert(context)

    assert result.success, result.error
    assert len(result.artifacts) == 1
    output = result.artifacts[0].staging_path
    with ZipFile(BytesIO(original)) as before, ZipFile(output) as after:
        old_style = _style(etree.fromstring(before.read("word/styles.xml")), _BODY_STYLE_ID)
        new_style = _style(etree.fromstring(after.read("word/styles.xml")), _BODY_STYLE_ID)
        assert _xml(new_style) == _xml(old_style)
        assert new_style.find(qn("w:aliases")) is not None
        for part in ("word/document.xml", "word/header1.xml", "word/footer1.xml"):
            expected = _field_paragraphs(etree.fromstring(before.read(part)))
            assert expected
            assert _field_paragraphs(etree.fromstring(after.read(part))) == expected
    reopened = Document(output)
    generated = next(paragraph for paragraph in reopened.paragraphs if paragraph.text == "New generated body.")
    assert generated.style is not None
    assert generated.style.style_id == _BODY_STYLE_ID
    assert template.read_bytes() == original


@pytest.mark.parametrize(
    "damage", ["name", "name_duplicate", "name_attributes", "aliases_changed", "aliases_removed", "style_id"]
)
def test_final_identity_audit_rejects_changes_to_preserved_template_identity(tmp_path: Path, damage: str) -> None:
    catalog = _catalog("en_US")
    completed, bindings = complete_managed_styles(_template("de_DE"), catalog)
    assert bindings.style_id("body_paragraph") == _BODY_STYLE_ID
    output = tmp_path / "identity-audit.docx"
    completed.save(str(output))
    validate_managed_style_package(output.read_bytes(), catalog, bindings)
    style = _style(completed.styles.element, _BODY_STYLE_ID)
    if damage.startswith("name"):
        name = style.find(qn("w:name"))
        assert name is not None
        if damage == "name":
            name.set(qn("w:val"), catalog.name_for("body_paragraph"))
        elif damage == "name_duplicate":
            style.append(deepcopy(name))
        else:
            name.set(qn("w:custom"), "changed")
    elif damage == "style_id":
        style.set(qn("w:styleId"), "RenamedAfterRendering")
    else:
        aliases = style.find(qn("w:aliases"))
        assert aliases is not None
        if damage == "aliases_removed":
            style.remove(aliases)
        else:
            aliases.set(qn("w:val"), "Changed alias")
    completed.save(str(output))

    with pytest.raises(ManagedStyleCompletionError, match="changed its completed identity"):
        validate_managed_style_package(output.read_bytes(), catalog, bindings)
