"""Reference text formatting preserves exact field and source authority."""

from __future__ import annotations

import hashlib
from collections.abc import Callable
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from lxml import etree

from docwen_core._docx_semantics_v3_model import CaptionStyleBindingV3, CaptionStyleKeyV3
from docwen_core._docx_semantics_v3_ooxml import inline_reference_sdt as v3_inline_reference_sdt
from docwen_core._docx_semantics_v3_ooxml import inline_sdt, recover_paragraph_children, sdt_tag
from docwen_core._docx_semantics_v3_topology import (
    prove_alias_run,
    prove_exact_field_run,
    prove_soft_reference_envelope,
)
from docwen_core.docx_host_metadata import field_run_payload
from docwen_core.docx_numbering_ooxml import inline_reference_sdt as v4_inline_reference_sdt
from docwen_core.docx_resolved_numbering import ResolvedNumberingDocxError, ResolvedNumberingDocxSession
from docwen_core.docx_resolved_numbering_recovery import ResolvedNumberingV4Recovery
from docwen_core.docx_semantics import DocxSemanticRenderer
from docwen_core.docx_semantics_v3 import (
    REFERENCE_OCCURRENCE_TAG_PREFIX,
    SOFT_REFERENCE_TAG_PREFIX,
    DocxSemanticsV3Error,
    DocxSemanticsV3Recovery,
    DocxSemanticsV3Session,
)
from docwen_core.models import resolved_numbering as numbering
from docwen_core.models.semantic_document import SemanticReference

pytestmark = pytest.mark.contract

_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
_FORMAT = (
    '<w:rPr><w:rStyle w:val="Emphasis"/><w:rFonts w:ascii="Aptos" w:eastAsia="SimSun" '
    'w:asciiTheme="minorHAnsi"/><w:b/><w:i w:val="1"/><w:color w:val="4F81BD" '
    'w:themeColor="accent1"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:vanish w:val="0"/></w:rPr>'
)


def _run(payload: str = "<w:t>1</w:t>", *, properties: str = _FORMAT) -> Any:
    return parse_xml(f"<w:r {_NS}>{properties}{payload}</w:r>")


@pytest.mark.parametrize("carrier", ["cache", "alias", "soft"])
def test_visual_properties_are_accepted_only_as_reference_text_formatting(carrier: str) -> None:
    run = _run('<w:t xml:space="preserve"> Alias</w:t>' if carrier == "alias" else "<w:t>1</w:t>")
    original = etree.tostring(run)
    assert field_run_payload(run) is None
    assert field_run_payload(run, allow_text_formatting=True) == [run.find(qn("w:t"))]

    if carrier == "cache":
        prove_exact_field_run(run, qn("w:t"))
    elif carrier == "alias":
        prove_alias_run(run)
    else:
        sdt = inline_sdt("owned", "1")
        content = sdt.find(qn("w:sdtContent"))
        content.replace(content[0], run)
        prove_soft_reference_envelope(sdt, "owned", "1")

    assert etree.tostring(run) == original


@pytest.mark.parametrize(
    "contents",
    [
        _FORMAT + _FORMAT + "<w:t>1</w:t>",
        "<w:t>1</w:t>" + _FORMAT,
        '<w:rPr w:unknown="1"><w:b/></w:rPr><w:t>1</w:t>',
        '<w:rPr><w:b w:unknown="1"/></w:rPr><w:t>1</w:t>',
        "<w:rPr><w:unknown/></w:rPr><w:t>1</w:t>",
        '<w:rPr><w:fldChar w:fldCharType="begin"/></w:rPr><w:t>1</w:t>',
        "<w:rPr><w:drawing/></w:rPr><w:t>1</w:t>",
        "<w:rPr><w:t>hidden</w:t></w:rPr><w:t>1</w:t>",
        "<w:rPr><w:rPrChange/></w:rPr><w:t>1</w:t>",
        '<w:rPr><w:rFonts w:asciiTheme="unknown"/></w:rPr><w:t>1</w:t>',
        '<w:rPr><w:sz w:val="-22"/></w:rPr><w:t>1</w:t>',
        '<w:rPr><w:color w:val="red"/></w:rPr><w:t>1</w:t>',
        "<w:rPr><w:vanish/></w:rPr><w:t>1</w:t>",
        "<w:rPr><w:webHidden/></w:rPr><w:t>1</w:t>",
        "<w:rPr><w:specVanish/></w:rPr><w:t>1</w:t>",
        "<w:rPr><w:oMath/></w:rPr><w:t>1</w:t>",
        _FORMAT + "<w:t>1</w:t><w:t>hidden</w:t>",
    ],
)
def test_text_formatting_does_not_admit_ambiguous_properties_or_hidden_payload(contents: str) -> None:
    run = parse_xml(f"<w:r {_NS}>{contents}</w:r>")
    with pytest.raises(DocxSemanticsV3Error, match="canonical"):
        prove_exact_field_run(run, qn("w:t"))


@pytest.mark.parametrize("payload_kind", ["marker", "instruction"])
def test_visual_properties_never_apply_to_field_markers_or_instructions(payload_kind: str) -> None:
    payload = (
        '<w:fldChar w:fldCharType="begin"/>'
        if payload_kind == "marker"
        else '<w:instrText xml:space="preserve"> REF foreign \\h </w:instrText>'
    )
    run = _run(payload)
    assert field_run_payload(run, allow_text_formatting=True) is None
    with pytest.raises(DocxSemanticsV3Error, match="canonical"):
        if payload_kind == "marker":
            prove_exact_field_run(run, qn("w:fldChar"), field_type="begin")
        else:
            prove_exact_field_run(run, qn("w:instrText"), xml_space="preserve")


@pytest.mark.parametrize("carrier", ["cache", "alias", "soft"])
def test_text_formatting_does_not_relax_exact_spacing(carrier: str) -> None:
    run = _run('<w:t xml:space="preserve">1</w:t>' if carrier == "cache" else "<w:t> Alias</w:t>")
    with pytest.raises(DocxSemanticsV3Error):
        if carrier == "cache":
            prove_exact_field_run(run, qn("w:t"))
        elif carrier == "alias":
            prove_alias_run(run)
        else:
            sdt = inline_sdt("owned", " Alias")
            content = sdt.find(qn("w:sdtContent"))
            content.replace(content[0], run)
            prove_soft_reference_envelope(sdt, "owned", " Alias")


def _heading_reference_port() -> numbering.ResolvedNumberingPort:
    token = "@[[#^intro|Alias]]"
    source = f"# Intro\n\n{token}\n"
    heading_end = source.index("\n")
    reference_start = source.index(token)
    source_sha = hashlib.sha256(source.encode()).hexdigest()
    heading = numbering.ResolvedDocumentTarget(
        source_start=0,
        source_end=heading_end,
        source_slice_sha256=hashlib.sha256(source[:heading_end].encode()).hexdigest(),
        kind="heading",
        target_id="intro",
        heading_level=1,
        authored_text="Intro",
    )
    reference = numbering.ResolvedReference(
        source_start=reference_start,
        source_end=reference_start + len(token),
        source_slice_sha256=hashlib.sha256(token.encode()).hexdigest(),
        authored_token=token,
        target_source_start=0,
        target_source_end=heading_end,
        target_kind="heading",
        target_id="intro",
        cached_number="1",
        alias="Alias",
    )
    definition = numbering.HeadingDefinition(
        "main",
        (
            numbering.HeadingLevelDefinition(
                1, 1, "arabic_half", (numbering.HeadingCounterSegment(1, "arabic_half"),), "space", None
            ),
        ),
    )
    plan = numbering.ResolvedNumberingPlan(
        (definition,),
        (numbering.HeadingInstance("document", "main", ()),),
        (
            numbering.NumberingTarget(
                0,
                heading_end,
                "heading",
                True,
                "intro",
                "1",
                numbering.HeadingListMaterialization("main", "document", 1),
            ),
        ),
    )
    return numbering.ResolvedNumberingPort(
        numbering.ResolvedDocumentEnvelope(
            "source", source_sha, "a" * 64, numbering.ResolvedDocument(source, (heading,), (reference,), (), (), ())
        ),
        numbering.NumberingExportPlanEnvelope("source", source_sha, "a" * 64, plan),
    )


def _caption_bindings(document: Any) -> tuple[CaptionStyleBindingV3, ...]:
    keys: tuple[CaptionStyleKeyV3, ...] = ("figure_caption", "table_caption", "equation_caption", "code_block_caption")
    bindings = []
    for index, key in enumerate(keys):
        style = document.styles.add_style(f"ReferenceFixtureCaption{index}", WD_STYLE_TYPE.PARAGRAPH)
        bindings.append(CaptionStyleBindingV3(key, style.style_id, style.name))
    return tuple(bindings)


def _mutate_document(path: Path, mutate: Callable[[bytes], bytes]) -> None:
    rewritten = path.with_suffix(".rewrite.docx")
    with ZipFile(path) as source, ZipFile(rewritten, "w") as output:
        for info in source.infolist():
            payload = source.read(info.filename)
            output.writestr(info, mutate(payload) if info.filename == "word/document.xml" else payload)
    rewritten.replace(path)


def _reference_package(tmp_path: Path, route: str) -> tuple[Any, Path]:
    document = Document()
    if route == "v3":
        stable, soft = "@[[#^intro|Alias]]", "@[[#Other|Soft]]"
        source = f"# Intro ^intro\n\n# Other\n\n{stable} {soft}\n"
        session: Any = DocxSemanticsV3Session(document, source_sha256=hashlib.sha256(source.encode()).hexdigest())
        session.bind_heading(document.add_heading("Intro", level=1), {"kind": "heading", "id": "intro"})
        document.add_heading("Other", level=1)
        paragraph = document.add_paragraph()
        for token, selector, alias, number in (
            (stable, "stable_id", "Alias", "1"),
            (soft, "heading_path", "Soft", "2"),
        ):
            start = source.index(token)
            session.render_reference(
                paragraph,
                {
                    "selector_kind": selector,
                    "resolution_status": "resolved",
                    "resolved_kind": "heading",
                    "target_id": "intro" if selector == "stable_id" else None,
                    "heading_path": ["Other"],
                    "alias": alias,
                    "cached_number": number,
                    "raw": token,
                    "range": {"start": start, "end": start + len(token)},
                },
            )
            paragraph.add_run(" ")
    else:
        session = ResolvedNumberingDocxSession(
            document,
            _heading_reference_port(),
            heading_style_ids={1: "Heading1"},
            heading_style_names={"heading_1": "Heading 1"},
            caption_style_bindings=_caption_bindings(document),
        )
        [heading_target] = session.port.document.targets
        session.bind_heading(
            document.add_heading("Intro", level=1),
            source_start=heading_target.source_start,
            source_end=heading_target.source_end,
        )
        paragraph = document.add_paragraph()
        reference = session.port.document.references[0]
        session.render_reference(paragraph, source_start=reference.source_start, source_end=reference.source_end)
    session.finalize_document()
    output = tmp_path / f"formatted-reference-{route}.docx"
    document.save(str(output))
    session.write_package(output)

    def format_text(data: bytes) -> bytes:
        root = etree.fromstring(data)
        for sdt in root.iter(qn("w:sdt")):
            if (sdt_tag(sdt) or "").startswith((REFERENCE_OCCURRENCE_TAG_PREFIX, SOFT_REFERENCE_TAG_PREFIX)):
                for text in sdt.iter(qn("w:t")):
                    run = text.getparent()
                    assert run is not None
                    run.insert(0, parse_xml(_FORMAT.replace("<w:rPr>", f"<w:rPr {_NS}>", 1)))
        return etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)

    _mutate_document(output, format_text)
    return session, output


def _recover(output: Path, route: str) -> DocxSemanticsV3Recovery:
    document = Document(str(output))
    if route == "v3":
        return DocxSemanticsV3Recovery.load(output, document)
    recovery = ResolvedNumberingV4Recovery.load_if_present(output, document)
    assert recovery is not None
    return recovery


@pytest.mark.parametrize("route", ["v3", "v4"])
def test_saved_formatted_reference_text_keeps_writer_and_fresh_reader_authority(tmp_path: Path, route: str) -> None:
    session, output = _reference_package(tmp_path, route)
    original = output.read_bytes()

    session.prove_package(output)
    recovery = _recover(output, route)

    assert recovery.reference_occurrence_identities[0].cached_number == "1"
    assert "|" in recovery.reference_occurrence_identities[0].authored_token
    if route == "v3":
        assert recovery.soft_reference_identities[0].authored_token == "@[[#Other|Soft]]"
    assert output.read_bytes() == original


@pytest.mark.parametrize("route", ["v3", "v4"])
@pytest.mark.parametrize("heading_number_only", [False, True])
def test_reference_writers_preserve_result_formatting_without_styling_field_identity(
    route: str, heading_number_only: bool
) -> None:
    bookmark = "DW_T_" + "a" * 35
    writer = v3_inline_reference_sdt if route == "v3" else v4_inline_reference_sdt
    control = writer(
        "owned",
        bookmark_name=bookmark,
        cached_number="1",
        heading_number_only=heading_number_only,
        alias="Alias",
    )

    [instruction] = list(control.iter(qn("w:instrText")))
    number_switch = " \\n" if heading_number_only else ""
    assert instruction.text == f" REF {bookmark}{number_switch} \\h \\* MERGEFORMAT "
    assert [node.text for node in control.iter(qn("w:t"))] == ["1", " Alias"]
    for run in control.iter(qn("w:r")):
        assert run.find(qn("w:rPr")) is None


def test_neutral_reference_writer_keeps_format_switch_after_save(tmp_path: Path) -> None:
    document = Document()
    renderer = DocxSemanticRenderer(document)
    target = SemanticReference("reference-target", "1")
    paragraph = document.add_paragraph()
    renderer.render_reference(paragraph, target)
    [instruction] = list(paragraph._p.iter(qn("w:instrText")))
    assert (instruction.text or "").endswith(" \\h \\* MERGEFORMAT ")
    output = tmp_path / "neutral-reference.docx"
    document.save(str(output))
    assert [(node.text or "") for node in Document(str(output)).element.iter(qn("w:instrText"))] == [instruction.text]
    # A foreign bookmark remains ordinary visible content on the v3 path.
    assert recover_paragraph_children(list(paragraph._p), target_ids_by_bookmark={}, soft_tokens_by_tag={}) == (
        "1",
        False,
    )


@pytest.mark.parametrize("route", ["v3", "v4"])
@pytest.mark.parametrize("syntax", ["legacy", "native-normalized"])
def test_reference_readers_accept_legacy_and_format_retaining_instruction_forms(
    tmp_path: Path, route: str, syntax: str
) -> None:
    session, output = _reference_package(tmp_path, route)

    def rewrite(data: bytes) -> bytes:
        root = etree.fromstring(data)
        [instruction] = list(root.iter(qn("w:instrText")))
        assert (instruction.text or "").endswith(" \\h \\* MERGEFORMAT ")
        if syntax == "legacy":
            instruction.text = (instruction.text or "").replace(" \\* MERGEFORMAT", "")
        else:
            # Horizontal whitespace and keyword case do not alter REF semantics.
            words = (instruction.text or "").split()
            instruction.text = "\t".join(word.lower() if index != 1 else word for index, word in enumerate(words))
        return etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)

    _mutate_document(output, rewrite)
    original = output.read_bytes()
    session.prove_package(output)
    recovery = _recover(output, route)
    assert recovery.reference_occurrence_identities[0].cached_number == "1"
    assert output.read_bytes() == original


@pytest.mark.parametrize("route", ["v3", "v4"])
@pytest.mark.parametrize(
    "damage",
    [
        "foreign-target",
        "caption-switches",
        "missing-hyperlink",
        "charformat",
        "numeric-format",
        "duplicate-format",
        "unknown-switch",
        "trailing-operand",
    ],
)
def test_format_retention_never_relaxes_reference_instruction_authority(
    tmp_path: Path, route: str, damage: str
) -> None:
    session, output = _reference_package(tmp_path, route)

    def rewrite(data: bytes) -> bytes:
        root = etree.fromstring(data)
        [instruction] = list(root.iter(qn("w:instrText")))
        value = instruction.text or ""
        assert value.endswith(" \\h \\* MERGEFORMAT ")
        if damage == "foreign-target":
            words = value.split()
            words[1] = "DW_T_" + "0" * 35
            value = " " + " ".join(words) + " "
        elif damage == "caption-switches":
            value = value.replace(" \\n", "")
        elif damage == "missing-hyperlink":
            value = value.replace(" \\h", "")
        elif damage == "charformat":
            value = value.replace("MERGEFORMAT", "CHARFORMAT")
        elif damage == "numeric-format":
            value += " \\* ROMAN "
        elif damage == "duplicate-format":
            value += " \\* MERGEFORMAT "
        elif damage == "unknown-switch":
            value += " \\p "
        else:
            value += " injected "
        instruction.text = value
        return etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)

    _mutate_document(output, rewrite)
    with pytest.raises((DocxSemanticsV3Error, ResolvedNumberingDocxError)):
        session.prove_package(output)
    with pytest.raises((DocxSemanticsV3Error, ResolvedNumberingDocxError)):
        _recover(output, route)


@pytest.mark.parametrize(
    ("route", "damage"),
    [("v3", "cache"), ("v4", "alias"), ("v4", "instruction"), ("v3", "tag"), ("v3", "soft_text")],
)
def test_formatted_reference_text_cannot_replace_authenticated_content_or_scope(
    tmp_path: Path, route: str, damage: str
) -> None:
    session, output = _reference_package(tmp_path, route)

    def mutate(data: bytes) -> bytes:
        root = etree.fromstring(data)
        prefix = SOFT_REFERENCE_TAG_PREFIX if damage == "soft_text" else REFERENCE_OCCURRENCE_TAG_PREFIX
        sdt = next(item for item in root.iter(qn("w:sdt")) if (sdt_tag(item) or "").startswith(prefix))
        content = sdt.find(qn("w:sdtContent"))
        assert content is not None
        if damage == "tag":
            tag = sdt.find(f"{qn('w:sdtPr')}/{qn('w:tag')}")
            assert tag is not None
            tag.set(qn("w:val"), prefix + "0" * 32)
        elif damage == "instruction":
            instruction = content.find(f".//{qn('w:instrText')}")
            assert instruction is not None
            instruction.text = f" REF DW_T_{'0' * 35} \\n \\h "
        else:
            texts = list(content.iter(qn("w:t")))
            texts[1 if damage == "alias" else 0].text = " forged"
        return etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)

    _mutate_document(output, mutate)
    with pytest.raises((DocxSemanticsV3Error, ResolvedNumberingDocxError)):
        session.prove_package(output)
    with pytest.raises((DocxSemanticsV3Error, ResolvedNumberingDocxError)):
        _recover(output, route)
