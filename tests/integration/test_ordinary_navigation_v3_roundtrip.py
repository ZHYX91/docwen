"""Heading navigation must not alter frozen v3 anchor/reference recovery."""

from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

import pytest
from docx import Document
from docx.oxml.ns import qn

from docwen_core.docx_semantics_v3 import ANCHOR_TAG_PREFIX, SOFT_REFERENCE_MAP_NAMESPACE, TARGET_MAP_NAMESPACE
from tests.integration._round_trip_helper import docx_to_md, md_to_docx

pytestmark = pytest.mark.integration


def test_ordinary_navigation_coexists_with_frozen_v3_recovery(tmp_path: Path, round_trip_runtime) -> None:
    source = tmp_path / "mixed-navigation.md"
    original = (
        b"[Jump](#intro) and [[#Other|Other jump]].\n\n"
        b"# Intro ^intro\n\n# Other\n\nPlain block ^raw\n\n"
        b"Frozen @[[#^intro|Intro]] and @[[#Other|Other]].\n\n"
        b"Ordinary [[#^raw|Raw block]].\n"
    )
    source.write_bytes(original)
    output = md_to_docx(round_trip_runtime, source, tmp_path / "forward")
    with ZipFile(output) as package:
        document_xml = package.read("word/document.xml").decode()
        package_text = b"\n".join(package.read(name) for name in package.namelist()).decode(errors="ignore")
    assert 'w:anchor="DWNav_' in document_xml
    assert ANCHOR_TAG_PREFIX in document_xml
    assert TARGET_MAP_NAMESPACE in package_text
    assert SOFT_REFERENCE_MAP_NAMESPACE in package_text
    assert " REF DW_T_" not in document_xml
    assert "SEQ " not in document_xml

    recovered = docx_to_md(
        round_trip_runtime,
        output,
        tmp_path / "reverse",
        options={"to_md_keep_images": False, "to_md_enable_ocr": False},
    )
    assert "[Jump](#intro)" in recovered
    assert "[Other jump](#Other)" in recovered
    assert "# Intro ^intro" in recovered
    assert "Plain block ^raw" in recovered
    assert "@[[#^intro|Intro]]" in recovered
    assert "@[[#Other|Other]]" in recovered
    assert "docwen-navigation:" not in recovered
    assert source.read_bytes() == original


def test_embedded_provenance_guard_preserves_original_root_anchor_recovery(tmp_path: Path, round_trip_runtime) -> None:
    source = tmp_path / "root.md"
    child = tmp_path / "child.md"
    child_bytes = b"# Embedded\n\nChild body.\n"
    child.write_bytes(child_bytes)
    original = b"[Child](#embedded) and [Root](#root).\n\n# Root ^root\n\nPlain root ^plain-root\n\n![[child.md]]\n\nFrozen @[[#^root|Root]].\n"
    source.write_bytes(original)
    output = md_to_docx(round_trip_runtime, source, tmp_path / "forward")
    with ZipFile(output) as package:
        document_xml = package.read("word/document.xml").decode()
    assert 'w:anchor="DWNav_' not in document_xml
    assert ANCHOR_TAG_PREFIX in document_xml
    recovered = docx_to_md(
        round_trip_runtime,
        output,
        tmp_path / "reverse",
        options={"to_md_keep_images": False, "to_md_enable_ocr": False},
    )
    assert "# Root ^root" in recovered
    assert "Plain root ^plain-root" in recovered
    assert "@[[#^root|Root]]" in recovered
    assert "Child body." in recovered
    assert "docwen-navigation:" not in recovered
    assert source.read_bytes() == original
    assert child.read_bytes() == child_bytes


def test_ordinary_navigation_survives_native_numbering_review(tmp_path: Path) -> None:
    from docwen_bundle.runtime_factory import create_runtime_port
    from docwen_runtime.config import ConfigLoader

    config = ConfigLoader(user_dir=tmp_path / "configs")
    assert config.set_value("template_fill.gongwen_numbering_comments", True)
    runtime = create_runtime_port(config_loader=config)
    source = tmp_path / "reviewed.md"
    original = b"[Jump](#intro)\n\n# Intro\n"
    source.write_bytes(original)
    output = md_to_docx(
        runtime,
        source,
        tmp_path / "forward",
        options={"add_numbering": True, "heading_numbering_render_mode": "word_native"},
    )
    document = Document(str(output))
    [link] = [node for node in document.element.iter(qn("w:hyperlink")) if node.get(qn("w:anchor"))]
    assert link.get(qn("w:anchor")).startswith("DWNav_")
    heading = next(paragraph for paragraph in document.paragraphs if paragraph.text == "Intro")
    assert list(heading._p.iter(qn("w:commentRangeStart")))
    assert len(document.comments) == 1

    recovered = docx_to_md(
        runtime,
        output,
        tmp_path / "reverse",
        options={"to_md_keep_images": False, "to_md_enable_ocr": False},
    )
    assert "[Jump](#intro)" in recovered
    assert source.read_bytes() == original


@pytest.mark.parametrize("tamper", ["hint", "anchor", "duplicate_heading"])
def test_untrusted_navigation_changes_import_as_plain_text(tmp_path: Path, round_trip_runtime, tamper: str) -> None:
    source = tmp_path / "navigation.md"
    source.write_text("[Jump](#intro)\n\n# Intro\n", encoding="utf-8")
    output = md_to_docx(round_trip_runtime, source, tmp_path / "forward")
    document = Document(str(output))
    link = next(node for node in document.element.iter(qn("w:hyperlink")) if node.get(qn("w:anchor")))
    if tamper == "hint":
        link.set(qn("w:tooltip"), "#Wrong")
    elif tamper == "anchor":
        link.set(qn("w:anchor"), "DWNav_" + "0" * 32)
    else:
        document.add_heading("Intro", level=1)
    modified = tmp_path / "modified.docx"
    document.save(str(modified))
    recovered = docx_to_md(
        round_trip_runtime,
        modified,
        tmp_path / "reverse",
        options={"to_md_keep_images": False, "to_md_enable_ocr": False},
    )
    assert "Jump" in recovered
    assert "[Jump](" not in recovered
    assert "DWNav_" not in recovered
    assert "docwen-navigation:" not in recovered
