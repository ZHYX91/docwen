"""Caption authentication accepts authored Word identities without trusting aliases."""

from __future__ import annotations

from io import BytesIO
from pathlib import Path
from typing import TYPE_CHECKING
from zipfile import ZipFile

import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

from docwen_core._docx_semantics_v3_styles import (
    CAPTION_STYLE_KEYS_V3,
    caption_style_binding_map_xml,
    parse_caption_style_binding_map,
    prove_caption_style_registry,
)
from docwen_core.docx_semantics_v3 import CaptionStyleBindingV3, DocxSemanticsV3Error

if TYPE_CHECKING:
    from docx.document import Document as DocumentObject

pytestmark = pytest.mark.contract


def _bindings(style_id: str) -> tuple[CaptionStyleBindingV3, ...]:
    return tuple(
        CaptionStyleBindingV3(key, style_id if index == 1 else f"Caption{index}", f"本地题注 {index}")
        for index, key in enumerate(CAPTION_STYLE_KEYS_V3)
    )


def _document(bindings: tuple[CaptionStyleBindingV3, ...]) -> DocumentObject:
    document = Document()
    for binding in bindings:
        style = document.styles.add_style(binding.visible_name, WD_STYLE_TYPE.PARAGRAPH)
        element = style._element
        assert element is not None
        element.set(qn("w:styleId"), binding.resolved_style_id)
        aliases = OxmlElement("w:aliases")
        aliases.set(qn("w:val"), "Template alias,模板别名")
        element.insert(1, aliases)
    return document


@pytest.mark.parametrize("style_id", ["9402", "图表_题注", '表&"<甲>', "Id\twith\nXML\rspace", "样" * 300])
def test_caption_map_and_saved_registry_preserve_xml_string_style_ids(tmp_path: Path, style_id: str) -> None:
    bindings = _bindings(style_id)
    payload = caption_style_binding_map_xml(bindings)
    assert parse_caption_style_binding_map(etree.fromstring(payload)) == bindings
    document = _document(bindings)
    path = tmp_path / "authored-caption.docx"
    document.save(str(path))
    with ZipFile(path) as package:
        assert prove_caption_style_registry(package, bindings) == bindings


@pytest.mark.parametrize("style_id", ["", "Bad\x00", "Bad\x08", "Bad\x0b", "Bad\ud800", "Bad\ufffe", "Bad\uffff"])
def test_caption_map_rejects_empty_or_non_xml_style_ids(style_id: str) -> None:
    with pytest.raises(DocxSemanticsV3Error, match="nonempty XML string"):
        caption_style_binding_map_xml(_bindings(style_id))


@pytest.mark.parametrize("damage", [None, "name_changed_to_alias", "ambiguous_name", "old_id_wrong_type"])
def test_office_id_rewrite_with_aliases_still_requires_exact_unique_primary_name(damage: str | None) -> None:
    bindings = _bindings("9402")
    document = _document(bindings)
    table_style = document.styles[bindings[1].visible_name]
    element = table_style._element
    assert element is not None
    element.set(qn("w:styleId"), "Word重新编号")
    if damage == "name_changed_to_alias":
        aliases = element.find(qn("w:aliases"))
        assert aliases is not None
        aliases.set(qn("w:val"), bindings[1].visible_name)
        table_style.name = "A different primary name"
    elif damage == "ambiguous_name":
        other = document.styles.add_style("Other identity", WD_STYLE_TYPE.PARAGRAPH)
        other.name = bindings[1].visible_name
    elif damage == "old_id_wrong_type":
        other = document.styles.add_style("Old identity collision", WD_STYLE_TYPE.CHARACTER)
        other_element = other._element
        assert other_element is not None
        other_element.set(qn("w:styleId"), "9402")
    buffer = BytesIO()
    document.save(buffer)
    with ZipFile(BytesIO(buffer.getvalue())) as package:
        with pytest.raises(DocxSemanticsV3Error, match="one exact paragraph style"):
            prove_caption_style_registry(package, bindings)
        if damage is None:
            rebound = prove_caption_style_registry(package, bindings, allow_style_id_rewrite=True)
            assert rebound[1] == CaptionStyleBindingV3("table_caption", "Word重新编号", bindings[1].visible_name)
        else:
            with pytest.raises(DocxSemanticsV3Error, match="one exact paragraph style"):
                prove_caption_style_registry(package, bindings, allow_style_id_rewrite=True)
