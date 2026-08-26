"""Formula conversion utilities shared by DOCX and Markdown plugins."""

from __future__ import annotations

from docwen_core.formula.constants import (
    BRACKET_PAIRS,
    LATEX2MATHML_AVAILABLE,
    MATHML_NS,
    MATRIX_BRACKET_MAP,
    NARY_OP_MAP,
    OMML_NS,
    SYMBOL_MAP,
)
from docwen_core.formula.latex_mathml import latex_to_mathml
from docwen_core.formula.markdown import parse_latex_from_markdown
from docwen_core.formula.mathml_latex import mathml_to_latex
from docwen_core.formula.mathml_omml import mathml_to_omml
from docwen_core.formula.omml_mathml import omml_to_mathml


def is_formula_supported() -> bool:
    """Check if formula conversion functionality is available."""
    return LATEX2MATHML_AVAILABLE


__all__ = [
    # constants
    "BRACKET_PAIRS",
    "MATHML_NS",
    "MATRIX_BRACKET_MAP",
    "NARY_OP_MAP",
    "OMML_NS",
    "SYMBOL_MAP",
    # public API
    "is_formula_supported",
    "latex_to_mathml",
    "mathml_to_latex",
    "mathml_to_omml",
    "omml_to_mathml",
    "parse_latex_from_markdown",
]
