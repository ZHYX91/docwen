"""Fifteen authored formulas survive two complete production round trips."""

from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

import pytest
from docx import Document
from lxml import etree

from tests.integration._round_trip_helper import _primary_path, _run

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]

_FIXTURE = Path(__file__).resolve().parents[1] / "fixtures" / "markdown" / "formula_roundtrip.md"
_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_NS = {"m": _M, "w": _W}
_DISPLAY = [False] * 5 + [True] * 4 + [False] * 3 + [True, False, False]
_ATOMS = [
    "E=mc2",
    "E=mc2",
    "a2+b2=c2",
    "α+β",
    "∫01xdx",
    "a2+b2=c2",
    "f′x=limh→0fx+h−fxh",
    "abcdxy=ax+bycx+dy",
    "fx=1σ2πe−(x−μ)22σ2",
    "E=mc2",
    "x2",
    "E=mc2",
    "a2+b2=c2",
    "E=mc2",
    "E=mc2",
]


def _math_signature(node):
    """Compare mathematical structure, omitting only run presentation."""
    if etree.QName(node).localname in {"rPr", "ctrlPr"}:
        return None
    return (
        node.tag,
        tuple(sorted(node.attrib.items())),
        node.text if etree.QName(node).localname == "t" else None,
        tuple(signature for child in node if (signature := _math_signature(child)) is not None),
    )


def _assert_native(path: Path):
    with ZipFile(path) as package:
        root = etree.fromstring(package.read("word/document.xml"))
    formulas = root.findall(".//m:oMath", _NS)
    assert len(formulas) == 15
    assert [node.getparent().tag == f"{{{_M}}}oMathPara" for node in formulas] == _DISPLAY
    assert [
        "".join(
            leaf.get(f"{{{_M}}}val", "") if leaf.tag == f"{{{_M}}}chr" else leaf.text or ""
            for leaf in node.iter()
            if leaf.tag in {f"{{{_M}}}t", f"{{{_M}}}chr"}
        )
        for node in formulas
    ] == _ATOMS
    matrices = formulas[7].findall(".//m:m", _NS)
    assert [[len(row.findall("m:e", _NS)) for row in matrix.findall("m:mr", _NS)] for matrix in matrices] == [
        [2, 2],
        [1, 1],
        [1, 1],
    ]
    assert formulas[6].find(".//m:func/m:fName/m:limLow/m:lim", _NS) is not None
    assert formulas[6].find(".//m:func/m:e/m:f", _NS) is not None
    assert formulas[8].find(".//m:radPr/m:degHide", _NS).get(f"{{{_M}}}val") == "1"
    # The second nested item keeps its real Word number and a left-aligned
    # display equation. The first item's inline identity is independent.
    display = formulas[12].getparent()
    paragraph = display.getparent()
    assert paragraph.find("w:pPr/w:numPr", _NS) is not None
    assert paragraph.find("w:pPr/w:ind", _NS) is None
    assert paragraph.find("w:pPr/w:jc", _NS).get(f"{{{_W}}}val") == "left"
    assert display.find("m:oMathParaPr/m:jc", _NS).get(f"{{{_M}}}val") == "left"
    for run in root.findall(".//m:r", _NS):
        names = [etree.QName(child).text for child in run]
        if f"{{{_M}}}rPr" in names and f"{{{_W}}}rPr" in names:
            assert names.index(f"{{{_M}}}rPr") < names.index(f"{{{_W}}}rPr") < names.index(f"{{{_M}}}t")
    return [_math_signature(node) for node in formulas]


def test_formula_source_survives_two_complete_runtime_round_trips(round_trip_runtime, tmp_path: Path) -> None:
    source = tmp_path / "formula-source.md"
    authored = _FIXTURE.read_bytes()
    source.write_bytes(authored)
    first_signature = None
    for cycle in range(2):
        before = source.read_bytes()
        output_dir = tmp_path / f"cycle-{cycle}"
        forward = _run(
            round_trip_runtime,
            f"formula-forward-{cycle}",
            source,
            source_format="markdown",
            target_format="docx",
            output_dir=output_dir,
            options={"heading_merge_mode": "never", "add_numbering": False},
        )
        docx = _primary_path(forward)
        assert not [item for item in forward.diagnostics if item.level in {"warning", "error"}]
        signature = _assert_native(docx)
        if first_signature is None:
            first_signature = signature
        else:
            assert signature == first_signature
        returned = _run(
            round_trip_runtime,
            f"formula-return-{cycle}",
            docx,
            source_format="docx",
            target_format="md",
            output_dir=output_dir / "returned",
            options={"remove_numbering": False},
        )
        result_path = _primary_path(returned)
        assert not [item for item in returned.diagnostics if item.level in {"warning", "error"}]
        text = result_path.read_text(encoding="utf-8")
        assert "$$\n\\begin{bmatrix}" in text
        assert "$not_math$" in text and "$$not_math$$" in text
        assert r"\$5 and \$10" in text
        assert source.read_bytes() == before
        source = result_path
    assert (tmp_path / "formula-source.md").read_bytes() == authored


def test_matrix_formula_in_table_and_list_keeps_its_container_for_two_cycles(
    round_trip_runtime, tmp_path: Path
) -> None:
    source = tmp_path / "container-matrices.md"
    matrix = r"\begin{bmatrix}a & b \\ c & d\end{bmatrix}"
    source.write_text(
        f"| Equation |\n| --- |\n| $${matrix}$$ |\n\n"
        "1. $$\\begin{bmatrix}\n   a & b \\\\\n   c & d\n   \\end{bmatrix}$$\n",
        encoding="utf-8",
    )
    signatures = None
    for cycle in range(2):
        output_dir = tmp_path / f"container-cycle-{cycle}"
        forward = _run(
            round_trip_runtime,
            f"container-formula-forward-{cycle}",
            source,
            source_format="markdown",
            target_format="docx",
            output_dir=output_dir,
            options={"heading_merge_mode": "never", "add_numbering": False},
        )
        assert not [item for item in forward.diagnostics if item.level in {"warning", "error"}]
        docx = _primary_path(forward)
        with ZipFile(docx) as package:
            root = etree.fromstring(package.read("word/document.xml"))
        formulas = root.findall(".//m:oMath", _NS)
        assert len(formulas) == 2
        assert all(node.getparent().tag == f"{{{_M}}}oMathPara" for node in formulas)
        assert any(ancestor.tag == f"{{{_W}}}tc" for ancestor in formulas[0].iterancestors())
        assert formulas[1].getparent().getparent().find("w:pPr/w:numPr", _NS) is not None
        for formula in formulas:
            matrix_node = formula.find(".//m:m", _NS)
            assert matrix_node is not None
            assert [len(row.findall("m:e", _NS)) for row in matrix_node.findall("m:mr", _NS)] == [2, 2]
        current = [_math_signature(node) for node in formulas]
        if signatures is None:
            signatures = current
        else:
            assert current == signatures
        returned = _run(
            round_trip_runtime,
            f"container-formula-return-{cycle}",
            docx,
            source_format="docx",
            target_format="md",
            output_dir=output_dir / "returned",
            options={"remove_numbering": False},
        )
        assert not [item for item in returned.diagnostics if item.level in {"warning", "error"}]
        source = _primary_path(returned)


def test_native_adjacent_formulas_and_literal_dollars_digits_survive_two_cycles(
    round_trip_runtime, tmp_path: Path
) -> None:
    source = tmp_path / "adjacent-native-formulas.docx"
    document = Document()
    paragraph = document.add_paragraph("$")
    for token in ("E", "x_y"):
        paragraph._p.append(etree.fromstring(f'<m:oMath xmlns:m="{_M}"><m:r><m:t>{token}</m:t></m:r></m:oMath>'))
    paragraph.add_run("2")
    document.save(source)
    for cycle in range(2):
        output_dir = tmp_path / f"adjacent-cycle-{cycle}"
        returned = _run(
            round_trip_runtime,
            f"adjacent-formula-return-{cycle}",
            source,
            source_format="docx",
            target_format="md",
            output_dir=output_dir / "returned",
            options={"remove_numbering": False},
        )
        assert not [item for item in returned.diagnostics if item.level in {"warning", "error"}]
        markdown = _primary_path(returned)
        assert r"\$ $E$ $x\_y$ 2" in markdown.read_text(encoding="utf-8")
        forward = _run(
            round_trip_runtime,
            f"adjacent-formula-forward-{cycle}",
            markdown,
            source_format="markdown",
            target_format="docx",
            output_dir=output_dir,
            options={"heading_merge_mode": "never", "add_numbering": False},
        )
        assert not [item for item in forward.diagnostics if item.level in {"warning", "error"}]
        source = _primary_path(forward)
        with ZipFile(source) as package:
            root = etree.fromstring(package.read("word/document.xml"))
        formulas = root.findall(".//m:oMath", _NS)
        assert len(formulas) == 2
        assert all(node.getparent().tag == f"{{{_W}}}p" for node in formulas)
        assert ["".join(leaf.text or "" for leaf in node.findall(".//m:t", _NS)) for node in formulas] == ["E", "x_y"]
