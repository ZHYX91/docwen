"""Native heading numbers belong to generated paragraphs, not template styles."""

from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree
from tests.support.numbering import repository_numbering_registry

from docwen_core.text.numbering_word_adapter import translate_scheme
from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter
from docwen_plugin_markdown.to_docx.heading_numbering import (
    capture_heading_numbering_targets,
    write_heading_numbering_to_docx,
)

from .conftest import make_context, write_temp_md

pytestmark = pytest.mark.contract


def _numbering(num_id: str, *, level: str = "0") -> etree._Element:
    properties = OxmlElement("w:numPr")
    for tag, value in (("ilvl", level), ("numId", num_id)):
        child = OxmlElement(f"w:{tag}")
        child.set(qn("w:val"), value)
        properties.append(child)
    return properties


def _translation():
    return translate_scheme(
        {"level_1": {"format": "{1.arabic_half} "}, "level_2": {"format": "{1.arabic_half}.{2.arabic_half} "}}
    )


def _style_properties(style: etree._Element) -> etree._Element:
    properties = style.find(qn("w:pPr"))
    if properties is None:
        properties = OxmlElement("w:pPr")
        style.append(properties)
    return properties


def _parts(path: Path) -> dict[str, bytes]:
    with ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


@pytest.mark.parametrize("rtl", [False, True], ids=["ltr", "rtl"])
def test_conversion_preserves_fixed_headings_and_shared_numbering(tmp_path: Path, rtl: bool) -> None:
    template = Document()
    heading_style = template.styles["Heading 1"]
    properties = _style_properties(heading_style.element)
    properties.insert(2, _numbering("1"))  # After keepNext/keepLines, before spacing/outlineLvl.
    if rtl:
        properties.insert(3, OxmlElement("w:bidi"))
    template.add_heading("Fixed before", level=1)
    template.add_paragraph("{{body}}")
    template.add_heading("Fixed after", level=1)
    template.sections[0].header.paragraphs[0].style = "Heading 1"
    template.sections[0].header.paragraphs[0].text = "Fixed header"
    original_style = etree.tostring(heading_style.element, method="c14n", exclusive=True)
    template_path = tmp_path / "template.docx"
    template.save(str(template_path))
    source_parts = _parts(template_path)

    context, _ = make_context(
        write_temp_md("# Generated first\n\n## Generated second\n\n# Generated third\n"),
        target_format="docx",
        options={
            "template_name": str(template_path),
            "add_numbering": True,
            "heading_numbering_render_mode": "word_native",
            "numbering_scheme": "hierarchical_standard",
        },
        config_values={"template_fill": {"gongwen_numbering_text": False}},
        numbering_registry=repository_numbering_registry(),
    )
    result = MdToDocxConverter().convert(context)
    assert result.success, result.error
    output_path = Path(result.artifacts[0].staging_path)
    output = Document(str(output_path))
    saved_style = output.styles["Heading 1"]
    assert etree.tostring(saved_style.element, method="c14n", exclusive=True) == original_style
    paragraphs = {paragraph.text: paragraph for paragraph in output.paragraphs}
    for text in ("Fixed before", "Fixed after"):
        assert paragraphs[text]._p.find(f"{qn('w:pPr')}/{qn('w:numPr')}") is None
        assert paragraphs[text].style == saved_style
    new_ids = set()
    for text, level in (("Generated first", "0"), ("Generated second", "1"), ("Generated third", "0")):
        properties = paragraphs[text]._p.find(f"{qn('w:pPr')}/{qn('w:numPr')}")
        assert properties is not None
        num_id = properties.find(qn("w:numId"))
        ilvl = properties.find(qn("w:ilvl"))
        assert num_id is not None and ilvl is not None
        new_ids.add(num_id.get(qn("w:val")))
        assert ilvl.get(qn("w:val")) == level
    assert len(new_ids) == 1 and "1" not in new_ids
    result_parts = _parts(output_path)
    assert result_parts["word/header1.xml"] == source_parts["word/header1.xml"]
    numbering = etree.fromstring(result_parts["word/numbering.xml"])
    num = numbering.find(f'{qn("w:num")}[@{qn("w:numId")}="{next(iter(new_ids))}"]')
    assert num is not None
    abstract_id = num.find(qn("w:abstractNumId"))
    assert abstract_id is not None
    abstract = numbering.find(f'{qn("w:abstractNum")}[@{qn("w:abstractNumId")}="{abstract_id.get(qn("w:val"))}"]')
    assert abstract is not None
    assert not abstract.findall(f"{qn('w:lvl')}/{qn('w:pStyle')}")
    first_level = abstract.find(qn("w:lvl"))
    assert first_level is not None
    alignment = first_level.find(qn("w:lvlJc"))
    assert alignment is not None and alignment.get(qn("w:val")) == ("right" if rtl else "left")
    indent = first_level.find(f"{qn('w:pPr')}/{qn('w:ind')}")
    assert indent is not None and indent.get(qn("w:right") if rtl else qn("w:left")) is not None


def test_saved_writer_preserves_other_parts_and_reserves_all_existing_ids(tmp_path: Path) -> None:
    document = Document()
    fixed = document.add_heading("Fixed", 1)
    fixed._p.get_or_add_pPr().append(_numbering("20000"))  # Dangling body reference.
    document.sections[0].header.paragraphs[0]._p.get_or_add_pPr().append(_numbering("20001"))
    _style_properties(document.styles["Heading 1"].element).insert(2, _numbering("20002"))
    numbering = document.part.numbering_part.element
    unused_num = OxmlElement("w:num")
    unused_num.set(qn("w:numId"), "20003")
    abstract_ref = OxmlElement("w:abstractNumId")
    abstract_ref.set(qn("w:val"), "0")
    unused_num.append(abstract_ref)
    numbering.append(unused_num)
    orphan_abstract = OxmlElement("w:abstractNum")
    orphan_abstract.set(qn("w:abstractNumId"), "20004")
    numbering.insert(0, orphan_abstract)
    cleanup = OxmlElement("w:numIdMacAtCleanup")
    cleanup.set(qn("w:val"), "20")
    numbering.append(cleanup)
    generated = document.add_heading("Generated", 1)
    targets = capture_heading_numbering_targets(document, [(generated, 1)])
    path = tmp_path / "numbered.docx"
    document.save(str(path))
    before = _parts(path)
    write_heading_numbering_to_docx(str(path), _translation(), targets=targets)
    after = _parts(path)
    assert set(before) == set(after)
    for part in before.keys() - {"word/document.xml", "word/numbering.xml"}:
        assert after[part] == before[part], part
    saved = Document(str(path))
    num_id = saved.paragraphs[-1]._p.find(f"{qn('w:pPr')}/{qn('w:numPr')}/{qn('w:numId')}")
    assert num_id is not None
    assert num_id.get(qn("w:val")) not in {str(value) for value in range(20000, 20005)}
    assert etree.tostring(saved.paragraphs[0]._p, method="c14n", exclusive=True) == etree.tostring(
        fixed._p, method="c14n", exclusive=True
    )
    saved_numbering = etree.fromstring(after["word/numbering.xml"])
    assert saved_numbering[-1].tag == qn("w:numIdMacAtCleanup")
    assert [etree.QName(child).localname for child in saved_numbering][-3:] == ["num", "num", "numIdMacAtCleanup"]


@pytest.mark.parametrize("change", ["style", "count", "path"])
def test_stale_targets_fail_without_changing_saved_file(tmp_path: Path, change: str) -> None:
    document = Document()
    generated = document.add_heading("Generated", 1)
    targets = capture_heading_numbering_targets(document, [(generated, 1)])
    if change == "style":
        generated.style = "Heading 2"
    elif change == "count":
        document.add_paragraph("New paragraph")
    else:
        generated._p.addprevious(OxmlElement("w:tbl"))
    path = tmp_path / "stale.docx"
    document.save(str(path))
    original = path.read_bytes()
    with pytest.raises(ValueError, match=r"heading|paragraphs"):
        write_heading_numbering_to_docx(str(path), _translation(), targets=targets)
    assert path.read_bytes() == original


def test_no_generated_heading_leaves_package_bytes_unchanged(tmp_path: Path) -> None:
    document = Document()
    document.add_heading("Fixed template heading", 1)
    targets = capture_heading_numbering_targets(document, [])
    path = tmp_path / "empty.docx"
    document.save(str(path))
    original = path.read_bytes()
    write_heading_numbering_to_docx(str(path), _translation(), targets=targets)
    assert path.read_bytes() == original


@pytest.mark.parametrize("include_supported", [False, True], ids=["unsupported-only", "mixed-levels"])
def test_unsupported_generated_level_suppresses_only_its_inherited_numbering(
    tmp_path: Path, include_supported: bool
) -> None:
    document = Document()
    _style_properties(document.styles["Heading 3"].element).insert(2, _numbering("1"))
    document.add_heading("Fixed level 3", 3)
    targets = [(document.add_heading("Generated level 3", 3), 3)]
    if include_supported:
        targets.append((document.add_heading("Generated level 1", 1), 1))
    snapshot = capture_heading_numbering_targets(document, targets)
    path = tmp_path / "unsupported.docx"
    document.save(str(path))
    before = _parts(path)
    write_heading_numbering_to_docx(str(path), _translation(), targets=snapshot)
    after = _parts(path)
    assert after["word/styles.xml"] == before["word/styles.xml"]
    if not include_supported:
        assert after["word/numbering.xml"] == before["word/numbering.xml"]
    saved = Document(str(path))
    assert saved.paragraphs[0]._p.find(f"{qn('w:pPr')}/{qn('w:numPr')}") is None
    suppressed = saved.paragraphs[1]._p.find(f"{qn('w:pPr')}/{qn('w:numPr')}/{qn('w:numId')}")
    assert suppressed is not None and suppressed.get(qn("w:val")) == "0"


def test_native_heading_targets_survive_repeated_note_projection() -> None:
    context, _ = make_context(
        write_temp_md("# First[^n]\n\n# Second[^n]\n\n[^n]: Shared note.\n"),
        target_format="docx",
        options={
            "add_numbering": True,
            "heading_numbering_render_mode": "word_native",
            "numbering_scheme": "hierarchical_standard",
        },
        numbering_registry=repository_numbering_registry(),
    )
    result = MdToDocxConverter().convert(context)
    assert result.success, result.error
    document = Document(str(result.artifacts[0].staging_path))
    headings = [paragraph for paragraph in document.paragraphs if paragraph.text.startswith(("First", "Second"))]
    assert len(headings) == 2
    assert all(paragraph._p.find(f"{qn('w:pPr')}/{qn('w:numPr')}") is not None for paragraph in headings)


def test_target_capture_requires_unique_attached_paragraphs() -> None:
    document = Document()
    other = Document().add_heading("Other document", 1)
    with pytest.raises(ValueError, match="attached"):
        capture_heading_numbering_targets(document, [(other, 1)])
    generated = document.add_heading("Generated", 1)
    with pytest.raises(ValueError, match="more than once"):
        capture_heading_numbering_targets(document, [(generated, 1), (generated, 1)])


def test_native_numpr_follows_frame_and_precedes_spacing(tmp_path: Path) -> None:
    document = Document()
    generated = document.add_heading("Framed heading", 1)
    properties = generated._p.get_or_add_pPr()
    frame = OxmlElement("w:framePr")
    frame.set(qn("w:w"), "3600")
    properties.append(frame)
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:after"), "120")
    properties.append(spacing)
    targets = capture_heading_numbering_targets(document, [(generated, 1)])
    path = tmp_path / "framed.docx"
    document.save(str(path))
    write_heading_numbering_to_docx(str(path), _translation(), targets=targets)

    saved = Document(str(path)).paragraphs[0]._p.get_or_add_pPr()
    assert [etree.QName(child).localname for child in saved] == ["pStyle", "framePr", "numPr", "spacing"]
    saved_frame = saved.find(qn("w:framePr"))
    saved_spacing = saved.find(qn("w:spacing"))
    assert saved_frame is not None and saved_frame.get(qn("w:w")) == "3600"
    assert saved_spacing is not None and saved_spacing.get(qn("w:after")) == "120"
