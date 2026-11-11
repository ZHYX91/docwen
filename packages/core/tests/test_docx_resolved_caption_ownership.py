"""Generated resolved captions have explicit ownership independent of template styles."""

from __future__ import annotations

import hashlib
import uuid
from copy import deepcopy
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import pytest
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

from docwen_core._docx_semantics_v3_caption_groups import (
    PAIRED_CAPTION_MAP_NAMESPACE,
    PAIRED_CAPTION_TAG_PREFIX,
    derive_paired_caption_group,
)
from docwen_core._docx_semantics_v3_model import DocxSemanticsV3Error
from docwen_core._docx_semantics_v3_ooxml import sdt_tag
from docwen_core.docx_resolved_numbering import ResolvedNumberingDocxSession
from docwen_core.docx_resolved_numbering_recovery import ResolvedNumberingV4Recovery
from docwen_core.docx_semantics_v3 import DocxSemanticsV3Recovery
from docwen_core.models.resolved_numbering import NumberingTarget, ResolvedDocumentTarget

from .test_docx_resolved_numbering_consistency_v4 import _caption_bindings, _caption_materialization, _two_table_port
from .test_docx_resolved_source_carriers_v4 import _anchor, _port, _sha
from .test_docx_semantics_v3_postreview import _mutate_xml_member

pytestmark = pytest.mark.contract


def _read(path: Path):
    document = Document(str(path))
    return document, ResolvedNumberingV4Recovery.load_if_present(path, document) or DocxSemanticsV3Recovery.load(
        path, document
    )


def _rebind_pair_map_uuid(path: Path, member: str) -> None:
    with ZipFile(path) as package:
        digest = hashlib.sha256(package.read(member)).hexdigest()
    value = "{" + str(uuid.uuid5(uuid.NAMESPACE_URL, f"{PAIRED_CAPTION_MAP_NAMESPACE}\0{digest}")).upper() + "}"
    props_member = member.replace("customXml/item", "customXml/itemProps", 1)

    def mutate(root: Any) -> None:
        root.set(f"{{{etree.QName(root).namespace}}}itemID", value)

    _mutate_xml_member(path, props_member, mutate)


def _output(tmp_path: Path, state: str, *, template_field: bool = True) -> tuple[ResolvedNumberingDocxSession, Path]:
    document = Document()
    bindings = _caption_bindings(document)
    fixed = document.add_paragraph("Fixed template caption", style="Table Caption V4")
    if template_field:
        field = etree.SubElement(fixed._p, qn("w:fldSimple"))
        field.set(qn("w:instr"), " PAGE ")
        etree.SubElement(etree.SubElement(field, qn("w:r")), qn("w:t")).text = "7"
    document.add_table(rows=1, cols=1).cell(0, 0).text = "Fixed table"
    original = etree.tostring(fixed._p, method="c14n")
    port = _two_table_port(state)
    session = ResolvedNumberingDocxSession(
        document, port, heading_style_ids={}, heading_style_names={}, caption_style_bindings=bindings
    )
    for target in port.document.targets:
        caption = document.add_paragraph(style="Table Caption V4")
        table = document.add_table(rows=1, cols=1)
        table.cell(0, 0).text = target.authored_text
        session.bind_caption(
            caption, (table._tbl,), source_start=target.source_start, source_end=target.source_end, kind="table"
        )
    output = tmp_path / "resolved-template-captions.docx"
    session.write_package(output)
    reopened, recovery = _read(output)
    assert etree.tostring(reopened.paragraphs[0]._p, method="c14n") == original
    assert not recovery.is_caption_element(reopened.paragraphs[0]._p)
    return session, output


@pytest.mark.parametrize("state", ["enabled_idless", "disabled_idless", "id_bearing"])
@pytest.mark.parametrize("template_field", [False, True])
def test_resolved_generated_caption_inventory_excludes_fixed_template_content(
    tmp_path: Path, state: str, template_field: bool
) -> None:
    session, output = _output(tmp_path, state, template_field=template_field)
    session.prove_package(output)
    _document, recovery = _read(output)
    assert [(item.title, item.cached_number) for item in recovery.recovered_captions] == [
        ("One", "" if state == "disabled_idless" else "1"),
        ("Two", "" if state == "disabled_idless" else "2"),
    ]
    assert recovery.paired_caption_map is not None and recovery.paired_caption_map.caption_count == 2
    assert len(recovery.paired_caption_map.groups) == (2 if state == "enabled_idless" else 0)


@pytest.mark.parametrize("mutation", ["missing_seq", "unwrapped", "duplicate", "missing_map", "missing_pair"])
def test_resolved_generated_pair_damage_is_rejected_by_fresh_reader(tmp_path: Path, mutation: str) -> None:
    _session, output = _output(tmp_path, "enabled_idless")
    if mutation in {"missing_map", "missing_pair"}:
        with ZipFile(output) as package:
            member = next(
                name
                for name in package.namelist()
                if name.startswith("customXml/item")
                and name.endswith(".xml")
                and etree.QName(etree.fromstring(package.read(name))).namespace == PAIRED_CAPTION_MAP_NAMESPACE
            )

        def mutate_map(root: Any) -> None:
            if mutation == "missing_map":
                root.tag = "unrelatedMap"
            else:
                root.remove(root[0])

        _mutate_xml_member(output, member, mutate_map)
        if mutation == "missing_pair":
            _rebind_pair_map_uuid(output, member)
    else:

        def mutate_document(root: Any) -> None:
            pair = next(
                node for node in root.iter(qn("w:sdt")) if (sdt_tag(node) or "").startswith(PAIRED_CAPTION_TAG_PREFIX)
            )
            if mutation == "missing_seq":
                next(pair.iter(qn("w:instrText"))).text = " PAGE "
            elif mutation == "duplicate":
                pair.addnext(deepcopy(pair))
            else:
                parent = pair.getparent()
                index = parent.index(pair)
                for child in reversed(list(pair.find(qn("w:sdtContent")))):
                    parent.insert(index, child)
                parent.remove(pair)

        _mutate_xml_member(output, "word/document.xml", mutate_document)
    with pytest.raises(DocxSemanticsV3Error):
        _read(output)


@pytest.mark.parametrize("title", ["", "Nested"])
def test_nested_idless_pair_recovers_exact_hierarchy_and_raw_object_anchor(tmp_path: Path, title: str) -> None:
    source = f"> Table: {title}\n> |a|\n> |-|\n> |b| ^raw\n\n^outer\n"
    end = source.index("\n\n^outer")
    target = ResolvedDocumentTarget(0, end, _sha(source[:end]), "table", None, None, title)
    plan = NumberingTarget(0, end, "table", True, None, "1", _caption_materialization(1))
    document = Document()
    session = ResolvedNumberingDocxSession(
        document,
        _port(source, (target,), (plan,)),
        heading_style_ids={},
        heading_style_names={},
        caption_style_bindings=_caption_bindings(document),
    )
    caption = document.add_paragraph(style="Table Caption V4")
    table = document.add_table(rows=1, cols=1)
    table.cell(0, 0).text = "b"
    raw = _anchor(
        source,
        "raw",
        "table",
        source.index("> |a|"),
        end,
        placement="inline",
        container_path=[{"block_kind": "block_quote", "block_range": {"start": 0, "end": end}}],
    )
    session.bind_ordinary_anchor((table._tbl,), raw, direct_parent_source_id="outer")
    session.bind_caption(caption, (table._tbl,), source_start=0, source_end=end, kind="table")
    session.bind_ordinary_anchor(
        (caption._p, table._tbl), _anchor(source, "outer", "block_quote", 0, end, placement="post_block")
    )
    output = tmp_path / "nested-caption-pair.docx"
    session.write_package(output)
    reopened = Document(str(output))
    recovery = ResolvedNumberingV4Recovery.load_if_present(output, reopened)
    assert recovery is not None
    [recovered] = recovery.recovered_captions
    assert (recovered.title, recovered.cached_number) == (title, "1")
    raw_anchor = recovery.source_anchor(recovered.object_elements[0])
    assert raw_anchor is not None and raw_anchor.source_id == "raw"
    session.prove_package(output)

    def mutate(root: Any) -> None:
        pair = next(
            node for node in root.iter(qn("w:sdt")) if (sdt_tag(node) or "").startswith(PAIRED_CAPTION_TAG_PREFIX)
        )
        # An unknown wrapper must never receive the proven nested-block allowance.
        pair.find(f"{qn('w:sdtPr')}/{qn('w:tag')}").set(qn("w:val"), PAIRED_CAPTION_TAG_PREFIX + "0" * 64)

    _mutate_xml_member(output, "word/document.xml", mutate)
    with pytest.raises(DocxSemanticsV3Error):
        _read(output)


def test_mixed_caption_carriers_share_one_complete_physical_ordinal_inventory(tmp_path: Path) -> None:
    fragments = ("Table: One\n|a|\n|-|\n|1|", "Table: Two\n|a|\n|-|\n|2|", "Table: Three\n|a|\n|-|\n|3|", "Table: Four")
    source = "\n\n".join(fragments)
    targets = []
    plans = []
    start = 0
    for index, (fragment, title) in enumerate(zip(fragments, ("One", "Two", "Three", "Four"), strict=True)):
        end = start + len(fragment)
        target_id = "named" if index == 2 else None
        enabled = index != 1
        sequence = 1 if index == 0 else index
        targets.append(ResolvedDocumentTarget(start, end, _sha(fragment), "table", target_id, None, title))
        plans.append(
            NumberingTarget(
                start,
                end,
                "table",
                enabled,
                target_id,
                str(sequence) if enabled else None,
                _caption_materialization(sequence) if enabled else None,
            )
        )
        start = end + 2
    document = Document()
    session = ResolvedNumberingDocxSession(
        document,
        _port(source, tuple(targets), tuple(plans)),
        heading_style_ids={},
        heading_style_names={},
        caption_style_bindings=_caption_bindings(document),
    )
    for index, target in enumerate(targets):
        caption = document.add_paragraph(style="Table Caption V4")
        objects = () if index == 3 else (document.add_table(rows=1, cols=1)._tbl,)
        session.bind_caption(
            caption, objects, source_start=target.source_start, source_end=target.source_end, kind="table"
        )
    output = tmp_path / "mixed-caption-ownership.docx"
    session.write_package(output)
    _document, recovery = _read(output)
    assert [(item.title, item.source_id, item.cached_number) for item in recovery.recovered_captions] == [
        ("One", None, "1"),
        ("Two", None, ""),
        ("Three", "named", "2"),
        ("Four", None, "3"),
    ]
    assert recovery.paired_caption_map is not None
    assert recovery.paired_caption_map.caption_count == 4
    [pair] = recovery.paired_caption_map.groups
    assert pair.ordinal == 1
    swapped = derive_paired_caption_group(_sha(source), 4, "table")
    with ZipFile(output) as package:
        member = next(
            name
            for name in package.namelist()
            if name.startswith("customXml/item")
            and name.endswith(".xml")
            and etree.QName(etree.fromstring(package.read(name))).namespace == PAIRED_CAPTION_MAP_NAMESPACE
        )

    def mutate_map(root: Any) -> None:
        root[0].set("ordinal", "4")
        root[0].set("tag", swapped.tag)

    def mutate_wrapper(root: Any) -> None:
        tag = next(node for node in root.iter(qn("w:tag")) if node.get(qn("w:val")) == pair.tag)
        tag.set(qn("w:val"), swapped.tag)

    _mutate_xml_member(output, member, mutate_map)
    _rebind_pair_map_uuid(output, member)
    _mutate_xml_member(output, "word/document.xml", mutate_wrapper)
    with pytest.raises(DocxSemanticsV3Error, match="ordinal differs"):
        _read(output)
