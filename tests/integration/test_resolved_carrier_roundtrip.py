"""Native package inventories and real reader assertions for resolved carriers."""

from __future__ import annotations

import hashlib
from copy import deepcopy
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import pytest
from docx.oxml.ns import qn
from lxml import etree

from docwen_core.models.resolved_numbering import NumberingTarget, ResolvedDocumentTarget
from tests.integration._round_trip_helper import _primary_path, _run

from ._resolved_carrier_roundtrip_support import (
    read_document,
    sha_text,
    write_caption_document,
    write_resolved_document,
)

pytestmark = pytest.mark.integration


@pytest.mark.parametrize(("quoted", "outer"), [(False, False), (True, False), (True, True)])
@pytest.mark.parametrize("structural_tables", [False, True])
def test_disabled_caption_table_survives_the_real_reader(
    tmp_path: Path, outer: bool, quoted: bool, structural_tables: bool
) -> None:
    body = "Table: Cap\n\n|a|\n|---|\n|b|\n"
    source = "".join("> " + line + "\n" if line else ">\n" for line in body.splitlines()) if quoted else body
    end = len(source.rstrip("\n"))
    if outer:
        source = source.rstrip("\n") + "\n\n^outer\n"
    output = write_caption_document(tmp_path / "write", source, target_end=end, structural_tables=structural_tables)
    with ZipFile(output) as package:
        document = etree.fromstring(package.read("word/document.xml"))
    wrappers = [
        item
        for item in document.iter(qn("w:sdt"))
        if item.find(f"{qn('w:sdtPr')}/{qn('w:tag')}").get(qn("w:val")).startswith("docwen-numbering-occurrence-v1:")
    ]
    assert len(wrappers) == 1
    content = wrappers[0].find(qn("w:sdtContent"))
    assert [item.tag for item in content] == [qn("w:p"), qn("w:tbl")]
    assert [item.text for item in content[1].iter(qn("w:t"))] == ["a", "b"]
    assert not list(wrappers[0].iter(qn("w:bookmarkStart")))
    assert not list(wrappers[0].iter(qn("w:instrText")))
    result, markdown = read_document(tmp_path / "read", output)
    assert result.success, result.error
    assert "Table: Cap" in markdown
    assert "| a |" in markdown and "| b |" in markdown
    assert markdown.count("^outer") == int(outer)


@pytest.mark.parametrize("outer", [False, True])
def test_caption_fence_is_bound_after_legitimate_quote_decoration(tmp_path: Path, outer: bool) -> None:
    source = "> Table: Cap\n>\n> ````text\n> literal `code`\n> ````\n"
    end = len(source.rstrip("\n"))
    if outer:
        source = source.rstrip("\n") + "\n\n^outer\n"
    output = write_caption_document(tmp_path / "write", source, target_end=end)
    with ZipFile(output) as package:
        document = etree.fromstring(package.read("word/document.xml"))
    literal = next(item for item in document.iter(qn("w:p")) if "literal" in "".join(item.itertext()))
    assert "".join(item.text or "" for item in literal.iter(qn("w:t"))) == "literal `code`"
    assert literal.find(f"{qn('w:pPr')}/{qn('w:pBdr')}/{qn('w:left')}") is not None
    result, markdown = read_document(tmp_path / "read", output)
    assert result.success, result.error
    assert "Table: Cap" in markdown
    assert "> ````text\n> literal `code`\n> ````" in markdown
    assert markdown.count("^outer") == int(outer)


@pytest.mark.parametrize("standalone_caption", [False, True])
def test_standalone_caption_keeps_independent_anchor_and_fence_inventory(
    tmp_path: Path, standalone_caption: bool
) -> None:
    prefix = "Figure: Planned ^planned\n\n\n" if standalone_caption else ""
    source = prefix + "ordinary paragraph ^ordinary\n\n````text\nliteral\n````\n"
    output = write_caption_document(
        tmp_path / "write",
        source,
        target_end=len("Figure: Planned ^planned") if standalone_caption else None,
        kind="figure",
        title="Planned",
        target_id="planned",
        structural_tables=False,
    )
    with ZipFile(output) as package:
        document = etree.fromstring(package.read("word/document.xml"))
        tags = [node.get(qn("w:val")) for node in document.iter(qn("w:tag"))]
        maps = b"".join(
            package.read(name) for name in package.namelist() if name.startswith("customXml/") and name.endswith(".xml")
        )
    assert sum(tag.startswith("docwen-anchor-v1:") for tag in tags) == 1
    assert sum(tag.startswith("docwen-fenced-source-v1:") for tag in tags) == 1
    # Only the explicit caption target owns a Word bookmark. Ordinary source
    # anchors remain Markdown navigation/recovery metadata.
    assert len(list(document.iter(qn("w:bookmarkStart")))) == int(standalone_caption)
    assert b'source_id="ordinary"' in maps
    assert "^ordinary" not in "".join(node.text or "" for node in document.iter(qn("w:t")))
    result, markdown = read_document(tmp_path / "read", output)
    assert result.success, result.error
    assert "ordinary paragraph ^ordinary" in markdown
    assert "````text\nliteral\n````" in markdown
    if standalone_caption:
        assert "Figure: Planned ^planned" in markdown


@pytest.mark.parametrize(
    ("kind", "carrier", "native_kind"),
    [("table", "````text\nliteral\n````", "p"), ("figure", "|a|\n|---|\n|b|", "tbl")],
)
def test_caption_kind_does_not_restrict_the_authenticated_nested_native_carrier(
    tmp_path: Path, kind: str, carrier: str, native_kind: str
) -> None:
    body = f"{kind.title()}: Cap\n\n{carrier}\n\n^raw\n"
    quoted = "".join("> " + line + "\n" if line else ">\n" for line in body.splitlines())
    source = quoted.rstrip("\n") + "\n\n^outer\n"
    output = write_caption_document(tmp_path / "write", source, target_end=len(quoted.rstrip("\n")), kind=kind)
    with ZipFile(output) as package:
        document = etree.fromstring(package.read("word/document.xml"))
    outer = next(document.iter(qn("w:sdt")))
    occurrence = outer.find(f"{qn('w:sdtContent')}/{qn('w:sdt')}")
    assert occurrence is not None
    content = occurrence.find(qn("w:sdtContent"))
    assert content is not None and len(content) == 2
    raw = content[0 if kind == "figure" else 1]
    assert raw.tag == qn("w:sdt")
    assert raw.find(f"{qn('w:sdtContent')}/{qn('w:' + native_kind)}") is not None
    result, markdown = read_document(tmp_path / "read", output)
    assert result.success, result.error
    assert f"{kind.title()}: Cap" in markdown
    assert markdown.count("^raw") == markdown.count("^outer") == 1
    if native_kind == "p":
        assert "> ````text\n> literal\n> ````" in markdown
    else:
        assert "| a |" in markdown and "| b |" in markdown


@pytest.mark.parametrize(
    "mutation", ["unknown_parent", "non_block_parent", "duplicate", "third_slot", "swap_slots", "missing_map", "field"]
)
def test_nested_carrier_acceptance_still_rejects_unowned_or_tampered_packages(tmp_path: Path, mutation: str) -> None:
    source = "> Table: Cap\n>\n> |a|\n> |---|\n> |b|\n\n^outer\n"
    output = write_caption_document(tmp_path / "write", source, target_end=source.index("\n\n^outer"))
    before = hashlib.sha256(output.read_bytes()).hexdigest()
    with ZipFile(output) as package:
        members = {name: package.read(name) for name in package.namelist()}
    document = etree.fromstring(members["word/document.xml"])
    body = document.find(qn("w:body"))
    assert body is not None
    occurrence = next(
        item
        for item in document.iter(qn("w:sdt"))
        if (item.find(f"{qn('w:sdtPr')}/{qn('w:tag')}").get(qn("w:val")) or "").startswith(
            "docwen-numbering-occurrence-v1:"
        )
    )
    content = occurrence.find(qn("w:sdtContent"))
    assert content is not None
    parent = occurrence.getparent()
    assert parent is not None
    if mutation == "unknown_parent":
        unknown = etree.Element(qn("w:sdt"))
        properties = etree.SubElement(unknown, qn("w:sdtPr"))
        etree.SubElement(properties, qn("w:tag")).set(qn("w:val"), "unrelated")
        inner = etree.SubElement(unknown, qn("w:sdtContent"))
        parent.replace(occurrence, unknown)
        inner.append(occurrence)
    elif mutation == "non_block_parent":
        paragraph = etree.Element(qn("w:p"))
        parent.replace(occurrence, paragraph)
        paragraph.append(occurrence)
    elif mutation == "duplicate":
        body.insert(0, deepcopy(occurrence))
    elif mutation == "third_slot":
        content.append(etree.Element(qn("w:p")))
    elif mutation == "swap_slots":
        content.append(content[0])
    elif mutation == "missing_map":
        map_name = next(
            name
            for name, data in members.items()
            if name.startswith("customXml/item") and b"<documentNumberingOccurrenceMap " in data
        )
        members[map_name] = b"<unrelated/>"
    else:
        paragraph = etree.SubElement(parent, qn("w:p"))
        run = etree.SubElement(paragraph, qn("w:r"))
        etree.SubElement(run, qn("w:instrText")).text = " REF forged \\h "
    members["word/document.xml"] = etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)
    tampered = tmp_path / "tampered.docx"
    with ZipFile(tampered, "w", ZIP_DEFLATED) as package:
        for name, data in members.items():
            package.writestr(name, data)
    result, markdown = read_document(tmp_path / "read", tampered)
    assert not result.success
    assert not result.artifacts and not markdown
    assert hashlib.sha256(output.read_bytes()).hexdigest() == before


@pytest.mark.parametrize("route", ["v3", "v4"])
def test_target_structural_table_target_share_live_bookmark_reservations(
    tmp_path: Path, round_trip_runtime, route: str
) -> None:
    source = (
        "# Before ^before\n\n"
        "| Region | Sales | < |\n| Quarter | Q1 | Q2 |\n| --- || --- | --- |\n| North | 10 | 12 |\n\n"
        "# After ^after\n"
    )
    if route == "v3":
        markdown = tmp_path / "source.md"
        markdown.write_text(source, encoding="utf-8")
        output = _primary_path(
            _run(
                round_trip_runtime,
                "target-table-target",
                markdown,
                source_format="markdown",
                target_format="docx",
                output_dir=tmp_path / "write",
                options={"heading_merge_mode": "never"},
            )
        )
    else:
        targets = []
        plan_targets = []
        for title, target_id in (("Before", "before"), ("After", "after")):
            start = source.index("# " + title)
            end = source.index("\n", start)
            targets.append(
                ResolvedDocumentTarget(start, end, sha_text(source[start:end]), "heading", target_id, 1, title)
            )
            plan_targets.append(NumberingTarget(start, end, "heading", False, target_id, None, None))
        output = write_resolved_document(tmp_path / "write", source, tuple(targets), tuple(plan_targets))
    with ZipFile(output) as package:
        document = etree.fromstring(package.read("word/document.xml"))
    starts = tuple(document.iter(qn("w:bookmarkStart")))
    ends = tuple(document.iter(qn("w:bookmarkEnd")))
    assert len(starts) == len(ends) == 3
    assert len({int(item.get(qn("w:id"))) for item in starts}) == 3
    assert len({item.get(qn("w:name")) for item in starts if item.get(qn("w:name")).startswith("DW_T_")}) == 2
    assert len({item.get(qn("w:name")) for item in starts if item.get(qn("w:name")).startswith("_DWT_")}) == 1
    visible = [item.text for item in document.iter(qn("w:t"))]
    assert visible == ["Before", "Region", "Sales", "Quarter", "Q1", "Q2", "North", "10", "12", "After"]
    result, returned = read_document(tmp_path / "read", output)
    assert result.success, result.error
    assert "Before ^before" in returned and "After ^after" in returned
    assert "North" in returned and "10" in returned and "12" in returned


@pytest.mark.parametrize("target_id", [None, "caption"])
def test_enabled_caption_in_authenticated_quote_retains_fields_title_and_anchor(
    tmp_path: Path, target_id: str | None
) -> None:
    id_text = "" if target_id is None else f" ^{target_id}"
    source = f"> Table: Cap{id_text}\n>\n> |a|\n> |---|\n> |b|\n\n^outer\n"
    output = write_caption_document(
        tmp_path / "write", source, target_end=source.index("\n\n^outer"), target_id=target_id, enabled=True
    )
    with ZipFile(output) as package:
        native = etree.fromstring(package.read("word/document.xml"))
    assert [node.text for node in native.iter(qn("w:instrText"))] == [" SEQ Table \\* ARABIC "]
    assert len(list(native.iter(qn("w:bookmarkStart")))) == int(target_id is not None)
    result, returned = read_document(tmp_path / "read", output)
    assert result.success, result.error
    assert "Table: Cap" + id_text in returned
    assert returned.count("^outer") == 1
    assert "| a |" in returned and "| b |" in returned


@pytest.mark.parametrize("enabled", [False, True])
@pytest.mark.parametrize("target_id", [None, "caption"])
def test_standalone_caption_can_share_its_authenticated_quote_owner(
    tmp_path: Path, enabled: bool, target_id: str | None
) -> None:
    id_text = "" if target_id is None else f" ^{target_id}"
    source = f"> Figure: Planned{id_text}\n>\n>\n> Body\n\n^outer\n"
    output = write_caption_document(
        tmp_path / "write",
        source,
        target_end=source.index("\n"),
        kind="figure",
        title="Planned",
        target_id=target_id,
        enabled=enabled,
    )
    with ZipFile(output) as package:
        native = etree.fromstring(package.read("word/document.xml"))
    tags = [item.get(qn("w:val")) for item in native.iter(qn("w:tag"))]
    assert tags[0].startswith("docwen-anchor-v1:")
    assert tags[1].startswith("docwen-target-v1:" if target_id else "docwen-standalone-caption-v1:")
    assert len(list(native.iter(qn("w:instrText")))) == int(enabled)
    result, returned = read_document(tmp_path / "read", output)
    assert result.success, result.error
    assert "Figure: Planned" + id_text in returned
    assert "Body" in returned and returned.count("^outer") == 1


@pytest.mark.parametrize("structural_tables", [False, True])
def test_list_anchor_owns_caption_and_native_table_without_quote_special_cases(
    tmp_path: Path, structural_tables: bool
) -> None:
    source = "- Table: Cap\n\n  |a|\n  |---|\n  |b|\n\n^outer\n"
    output = write_caption_document(
        tmp_path / "write", source, target_end=source.index("\n\n^outer"), structural_tables=structural_tables
    )
    with ZipFile(output) as package:
        native = etree.fromstring(package.read("word/document.xml"))
    tags = [item.get(qn("w:val")) for item in native.iter(qn("w:tag"))]
    assert tags[0].startswith("docwen-anchor-v1:")
    assert tags[1].startswith("docwen-numbering-occurrence-v1:")
    assert [item.text for item in native.iter(qn("w:t"))] == ["Cap", "a", "b"]
    tables = list(native.iter(qn("w:tbl")))
    assert len(tables) == 1
    assert [
        ["".join(item.text or "" for item in cell.iter(qn("w:t"))) for cell in row.findall(qn("w:tc"))]
        for row in tables[0].findall(qn("w:tr"))
    ] == [["a"], ["b"]]
    result, returned = read_document(tmp_path / "read", output)
    assert result.success, result.error
    assert "Table: Cap" in returned and returned.count("^outer") == 1
    assert "| a |" in returned and "| b |" in returned


@pytest.mark.parametrize("mutation", ["unowned_bookmark", "wrong_seq"])
def test_enabled_idless_caption_candidate_requires_its_independent_field_proof(tmp_path: Path, mutation: str) -> None:
    source = "> Table: Cap\n>\n> |a|\n> |---|\n> |b|\n\n^outer\n"
    output = write_caption_document(tmp_path / "write", source, target_end=source.index("\n\n^outer"), enabled=True)
    original = hashlib.sha256(output.read_bytes()).hexdigest()
    with ZipFile(output) as package:
        members = {name: package.read(name) for name in package.namelist()}
    document = etree.fromstring(members["word/document.xml"])
    instruction = next(document.iter(qn("w:instrText")))
    assert instruction.text == " SEQ Table \\* ARABIC "
    if mutation == "wrong_seq":
        instruction.text = " SEQ Figure \\* ARABIC "
    else:
        caption = next(ancestor for ancestor in instruction.iterancestors() if ancestor.tag == qn("w:p"))
        start = etree.SubElement(caption, qn("w:bookmarkStart"))
        start.set(qn("w:name"), "Unowned")
        start.set(qn("w:id"), "99")
        etree.SubElement(caption, qn("w:bookmarkEnd")).set(qn("w:id"), "99")
    members["word/document.xml"] = etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)
    tampered = tmp_path / "tampered.docx"
    with ZipFile(tampered, "w", ZIP_DEFLATED) as package:
        for name, data in members.items():
            package.writestr(name, data)
    result, markdown = read_document(tmp_path / "read", tampered)
    assert not result.success and not result.artifacts and not markdown
    assert hashlib.sha256(output.read_bytes()).hexdigest() == original
