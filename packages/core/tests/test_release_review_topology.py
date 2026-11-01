"""Interval sweep ownership controls, including large disjoint inventories."""

from types import SimpleNamespace

import pytest
from docx import Document

from docwen_core._docx_semantics_v3_topology import order_ordinary_anchor_bindings
from docwen_core.docx_semantics_v3 import DocxSemanticsV3Error, derive_anchor_identity_v3

pytestmark = pytest.mark.contract


def _binding(name, elements, parent=None):
    return SimpleNamespace(
        identity=derive_anchor_identity_v3("block_quote", name),
        elements=tuple(elements),
        direct_parent_source_id=parent,
    )


def test_release_review_topology_large_disjoint_inventory() -> None:
    document = Document()
    paragraphs = [document.add_paragraph(str(index))._p for index in range(1500)]
    bindings = [_binding(f"b{index}", [paragraph]) for index, paragraph in enumerate(paragraphs)]
    ordered, edges = order_ordinary_anchor_bindings(document._element.body, bindings)
    assert ordered == bindings
    assert not edges


def test_release_review_topology_equal_nested_ranges() -> None:
    document = Document()
    paragraph = document.add_paragraph("nested")._p
    bindings = [_binding(f"b{index}", [paragraph], f"b{index - 1}" if index else None) for index in range(128)]
    ordered, edges = order_ordinary_anchor_bindings(document._element.body, list(reversed(bindings)))
    assert ordered == list(reversed(bindings))
    assert len(edges) == 127


@pytest.mark.parametrize(
    "case", ["crossing", "missing-parent", "disjoint-parent", "duplicate", "nested-without-parent"]
)
def test_release_review_topology_rejects_invalid_ownership(case: str) -> None:
    document = Document()
    blocks = [document.add_paragraph(str(index))._p for index in range(3)]
    if case == "crossing":
        bindings = [_binding("a", blocks[:2]), _binding("b", blocks[1:], "a")]
    elif case == "missing-parent":
        bindings = [_binding("a", blocks[:1], "missing")]
    elif case == "disjoint-parent":
        bindings = [_binding("a", blocks[:1]), _binding("b", blocks[1:], "a")]
    elif case == "duplicate":
        bindings = [_binding("a", blocks[:1]), _binding("a", blocks[1:])]
    else:
        bindings = [_binding("a", blocks), _binding("b", blocks[:1])]
    with pytest.raises(DocxSemanticsV3Error):
        order_ordinary_anchor_bindings(document._element.body, bindings)
