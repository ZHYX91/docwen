"""Source-ordered text and OMML formula extraction for Gongwen paragraphs."""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING, Any

from docwen_core.docx_parsing.xml_ns import NS_M, NS_W
from docwen_plugin_optimizer_gongwen.models import ParagraphAtom

if TYPE_CHECKING:
    from docx.text.paragraph import Paragraph

logger = logging.getLogger(__name__)


def extract_formula_atoms(para: Paragraph) -> tuple[ParagraphAtom, ...]:
    """Claim each outer math element once and retain surrounding text order."""
    atoms: list[ParagraphAtom] = []

    def visit(node: Any) -> None:
        if node.tag in {f"{{{NS_W}}}del", f"{{{NS_W}}}moveFrom"}:
            return
        if node.tag == f"{{{NS_W}}}r":
            vanish = node.find(f"{{{NS_W}}}rPr/{{{NS_W}}}vanish")
            if vanish is not None:
                value = vanish.get(f"{{{NS_W}}}val", "1").strip().casefold()
                if value not in {"0", "false", "off", "no", "none"}:
                    return
        if node.tag in {f"{{{NS_M}}}oMath", f"{{{NS_M}}}oMathPara"}:
            latex = _omml_element_to_latex(node)
            # The outer oMathPara owns its nested oMath; do not visit it again.
            if latex:
                kind = "block_math" if node.tag == f"{{{NS_M}}}oMathPara" else "inline_math"
                atoms.append(ParagraphAtom(kind, latex))
            return
        if node.tag == f"{{{NS_W}}}t":
            atoms.append(ParagraphAtom("text", node.text or ""))
            return
        if node.tag in {f"{{{NS_W}}}br", f"{{{NS_W}}}cr", f"{{{NS_W}}}tab"}:
            atoms.append(ParagraphAtom("text", "\t" if node.tag == f"{{{NS_W}}}tab" else "\n"))
            return
        # Images/textboxes have separate consumers; their descendant text is
        # not part of this paragraph's inline text/math ownership.
        if node.tag in {f"{{{NS_W}}}drawing", f"{{{NS_W}}}pict", f"{{{NS_W}}}object"}:
            return
        if node.tag == "{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent":
            children = list(node)
            choice = next((child for child in children if child.tag.endswith("}Choice")), None)
            fallback = next((child for child in children if child.tag.endswith("}Fallback")), None)
            selected = choice if choice is not None else fallback
            if selected is not None:
                visit(selected)
            return
        for child in node:
            visit(child)

    visit(para._p)
    return tuple(atoms)


def extract_formula_info(para: Paragraph) -> tuple[bool, str, str]:
    """Compatibility summary derived from complete structural occurrences."""
    return summarize_formula_atoms(extract_formula_atoms(para))


def summarize_formula_atoms(atoms: tuple[ParagraphAtom, ...]) -> tuple[bool, str, str]:
    math_atoms = [atom for atom in atoms if atom.kind != "text"]
    if not math_atoms:
        return False, "", ""
    has_prose = any(atom.kind == "text" and atom.text.strip() for atom in atoms)
    return True, "inline" if has_prose else "block", " ".join(atom.text for atom in math_atoms)


def formula_text(atoms: tuple[ParagraphAtom, ...]) -> str:
    """Return the same visible text projection used for structure detection."""
    return "".join(atom.text for atom in atoms if atom.kind == "text")


def strip_formula_prefix(atoms: tuple[ParagraphAtom, ...], prefix: str) -> tuple[ParagraphAtom, ...]:
    """Remove one detected source prefix, without crossing an owned formula."""
    if not prefix:
        return atoms
    leading: list[str] = []
    for atom in atoms:
        if atom.kind != "text":
            break
        leading.append(atom.text)
    if not "".join(leading).startswith(prefix):
        return atoms
    remaining = len(prefix)
    result: list[ParagraphAtom] = []
    for atom in atoms:
        if remaining and atom.kind == "text":
            consumed = min(len(atom.text), remaining)
            remaining -= consumed
            if consumed < len(atom.text):
                result.append(ParagraphAtom("text", atom.text[consumed:]))
        else:
            result.append(atom)
    return tuple(result)


def split_formula_atoms(
    atoms: tuple[ParagraphAtom, ...], text_offset: int, *, include_boundary_math: bool = False
) -> tuple[tuple[ParagraphAtom, ...], tuple[ParagraphAtom, ...]]:
    """Project a clean-text heading boundary without counting LaTeX markup.

    A direct run-format boundary starts at its following body run, so formulas
    preceding that run stay with the heading. A punctuation fallback starts
    immediately after the delimiter and does not claim following formulas.
    """
    text = formula_text(atoms)
    remaining = text_offset + len(text) - len(text.lstrip())
    left: list[ParagraphAtom] = []
    right: list[ParagraphAtom] = []
    after_boundary = False
    for atom in atoms:
        if after_boundary:
            right.append(atom)
        elif atom.kind != "text":
            if remaining == 0 and not include_boundary_math:
                right.append(atom)
                after_boundary = True
            else:
                left.append(atom)
        elif len(atom.text) <= remaining:
            left.append(atom)
            remaining -= len(atom.text)
        else:
            if remaining:
                left.append(ParagraphAtom("text", atom.text[:remaining]))
            right.append(ParagraphAtom("text", atom.text[remaining:]))
            remaining = 0
            after_boundary = True
    return tuple(left), tuple(right)


def render_formula_atoms(atoms: tuple[ParagraphAtom, ...]) -> str:
    """Render known source atoms without moving formulas past authored prose."""
    standalone = not any(atom.kind == "text" and atom.text.strip() for atom in atoms)
    parts: list[str] = []
    for atom in atoms:
        if atom.kind == "text":
            parts.append(atom.text)
        elif atom.kind == "block_math" or standalone:
            parts.append(f"\n\n$$\n{atom.text}\n$$\n\n")
        else:
            parts.append(f"${atom.text}$")
    return "".join(parts).strip()


def _omml_element_to_latex(om_elem: Any) -> str:
    import lxml.etree as etree

    from docwen_core.formula import mathml_to_latex, omml_to_mathml

    omml_xml = etree.tostring(om_elem, encoding="unicode")
    mathml = omml_to_mathml(omml_xml)
    if not mathml:
        return ""
    return mathml_to_latex(mathml) or ""
