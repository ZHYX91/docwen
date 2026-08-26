"""Tests for Word-native list numbering in MD→DOCX conversion.

Covers findings: F-F1-014, F-F1-015, F-F1-016, F-F1-017, F-F3-023
"""

from __future__ import annotations

import copy
import hashlib
import os
import tempfile
import zipfile
from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pytest
from lxml import etree

pytestmark = pytest.mark.contract

from tests.support.numbering import repository_numbering_registry

from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter
from docwen_plugin_markdown.to_docx.numbering import (
    DocxListNumbering,
    apply_list_to_paragraph,
    write_numbering_to_docx,
)

from .conftest import PROJECT_ROOT, make_context, write_temp_md

WML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

BUNDLED_DOCX_TEMPLATES = tuple(sorted((PROJECT_ROOT / "templates").glob("*.docx")))


class _ExactSchemeRegistry:
    def __init__(self, *, enabled: bool = True, levels: dict[str, str] | None = None) -> None:
        self._scheme = SimpleNamespace(
            enabled=enabled,
            levels={"level_1": "{1.arabic_half} "} if levels is None else levels,
        )

    def get_scheme(self, scheme_id: str) -> object:
        if scheme_id != "exact":
            raise LookupError(scheme_id)
        return self._scheme


def _get_paragraphs_xml(docx_path: str) -> list[etree._Element]:
    """Extract all ``w:p`` elements from ``word/document.xml``."""
    with zipfile.ZipFile(docx_path, "r") as zf:
        doc_xml = zf.read("word/document.xml")
    root = etree.fromstring(doc_xml)
    body = root.find(f"{{{WML_NS}}}body")
    assert body is not None, "Missing w:body in document.xml"
    return body.findall(f".//{{{WML_NS}}}p")


def _get_numbering_xml(docx_path: str) -> etree._Element | None:
    """Return ``numbering.xml`` root element, or None."""
    with zipfile.ZipFile(docx_path, "r") as zf:
        if "word/numbering.xml" not in zf.namelist():
            return None
        raw = zf.read("word/numbering.xml")
    return etree.fromstring(raw)


def _list_paras_with_numpr(docx_path: str) -> list[etree._Element]:
    """Return paragraphs that contain ``w:numPr``."""
    result: list[etree._Element] = []
    for p in _get_paragraphs_xml(docx_path):
        pPr = p.find(f"{{{WML_NS}}}pPr")
        if pPr is not None and pPr.find(f"{{{WML_NS}}}numPr") is not None:
            result.append(p)
    return result


def _our_abstract_num_id(num_root: etree._Element) -> str:
    """Return the abstractNumId referenced by the last ``w:num`` element.

    This is the numbering definition created by the converter during
    this run (since we use high IDs to avoid template conflicts).
    """
    nums = num_root.findall(f"{{{WML_NS}}}num")
    assert nums, "No w:num elements found in numbering.xml"
    last_num = nums[-1]
    an_ref = last_num.find(f"{{{WML_NS}}}abstractNumId")
    assert an_ref is not None
    val = an_ref.get(f"{{{WML_NS}}}val")
    assert val is not None, "Missing val attribute on abstractNumId"
    return val


def _append_template_abstract_num(document: Any, abstract_num_id: int, levels: list[dict[str, str]]) -> None:
    abstract_num = etree.Element(f"{{{WML_NS}}}abstractNum")
    abstract_num.set(f"{{{WML_NS}}}abstractNumId", str(abstract_num_id))
    for spec in levels:
        level = etree.SubElement(abstract_num, f"{{{WML_NS}}}lvl")
        level.set(f"{{{WML_NS}}}ilvl", spec["level"])
        if "tentative" in spec:
            level.set(f"{{{WML_NS}}}tentative", spec["tentative"])
        for tag, key in (("start", "start"), ("numFmt", "numFmt"), ("lvlText", "lvlText"), ("lvlJc", "lvlJc")):
            if key in spec:
                element = etree.SubElement(level, f"{{{WML_NS}}}{tag}")
                element.set(f"{{{WML_NS}}}val", spec[key])
        if "left" in spec or "hanging" in spec:
            p_pr = etree.SubElement(level, f"{{{WML_NS}}}pPr")
            indent = etree.SubElement(p_pr, f"{{{WML_NS}}}ind")
            if "left" in spec:
                indent.set(f"{{{WML_NS}}}left", spec["left"])
            if "hanging" in spec:
                indent.set(f"{{{WML_NS}}}hanging", spec["hanging"])
        if "ascii" in spec or "hAnsi" in spec:
            r_pr = etree.SubElement(level, f"{{{WML_NS}}}rPr")
            fonts = etree.SubElement(r_pr, f"{{{WML_NS}}}rFonts")
            if "ascii" in spec:
                fonts.set(f"{{{WML_NS}}}ascii", spec["ascii"])
            if "hAnsi" in spec:
                fonts.set(f"{{{WML_NS}}}hAnsi", spec["hAnsi"])

    numbering_root = document.part.numbering_part.element
    first_num = numbering_root.find(f"{{{WML_NS}}}num")
    if first_num is None:
        numbering_root.append(abstract_num)
    else:
        first_num.addprevious(abstract_num)


__all__ = (
    "BUNDLED_DOCX_TEMPLATES",
    "WML_NS",
    "DocxListNumbering",
    "MdToDocxConverter",
    "Path",
    "_ExactSchemeRegistry",
    "_append_template_abstract_num",
    "_get_numbering_xml",
    "_list_paras_with_numpr",
    "_our_abstract_num_id",
    "apply_list_to_paragraph",
    "copy",
    "etree",
    "hashlib",
    "make_context",
    "os",
    "pytest",
    "pytestmark",
    "repository_numbering_registry",
    "tempfile",
    "write_numbering_to_docx",
    "write_temp_md",
    "zipfile",
)
