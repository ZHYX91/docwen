"""Formula constants: namespaces, symbol maps, matrix brackets, LaTeX availability flag."""

from __future__ import annotations

import logging

try:
    from latex2mathml.converter import convert as _latex_to_mathml_convert  # type: ignore[unused-ignore]

    LATEX2MATHML_AVAILABLE = True
except ImportError:
    LATEX2MATHML_AVAILABLE = False
    _latex_to_mathml_convert = None  # type: ignore[assignment]
    logging.warning("latex2mathml not installed; formula conversion unavailable")

# Office Math namespace
OMML_NS: dict[str, str] = {
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}

# MathML namespace
MATHML_NS: str = "http://www.w3.org/1998/Math/MathML"

# Special symbol mapping (Unicode -> LaTeX)
SYMBOL_MAP: dict[str, str] = {
    "α": r"\alpha",
    "β": r"\beta",
    "γ": r"\gamma",
    "δ": r"\delta",
    "ε": r"\epsilon",
    "ζ": r"\zeta",
    "η": r"\eta",
    "θ": r"\theta",
    "ι": r"\iota",
    "κ": r"\kappa",
    "λ": r"\lambda",
    "μ": r"\mu",
    "ν": r"\nu",
    "ξ": r"\xi",
    "ο": r"o",
    "π": r"\pi",
    "ρ": r"\rho",
    "σ": r"\sigma",
    "τ": r"\tau",
    "υ": r"\upsilon",
    "φ": r"\phi",
    "χ": r"\chi",
    "ψ": r"\psi",
    "ω": r"\omega",
    "Α": r"A",
    "Β": r"B",
    "Γ": r"\Gamma",
    "Δ": r"\Delta",
    "Ε": r"E",
    "Ζ": r"Z",
    "Η": r"H",
    "Θ": r"\Theta",
    "Ι": r"I",
    "Κ": r"K",
    "Λ": r"\Lambda",
    "Μ": r"M",
    "Ν": r"N",
    "Ξ": r"\Xi",
    "Ο": r"O",
    "Π": r"\Pi",
    "Ρ": r"P",
    "Σ": r"\Sigma",
    "Τ": r"T",
    "Υ": r"\Upsilon",
    "Φ": r"\Phi",
    "Χ": r"X",
    "Ψ": r"\Psi",
    "Ω": r"\Omega",
    "∑": r"\sum",
    "∏": r"\prod",
    "∫": r"\int",
    "∬": r"\iint",
    "∮": r"\oint",
    "∂": r"\partial",
    "∇": r"\nabla",
    "∞": r"\infty",
    "lim": r"\lim",
    "→": r"\to",
    "←": r"\gets",
    "↔": r"\leftrightarrow",
    "⇒": r"\Rightarrow",
    "≈": r"\approx",
    "≠": r"\neq",
    "≤": r"\leq",
    "≥": r"\geq",
    "±": r"\pm",
    "×": r"\times",
    "÷": r"\div",
    "∈": r"\in",
    "∉": r"\notin",
    "⊂": r"\subset",
    "⊃": r"\supset",
    "∪": r"\cup",
    "∩": r"\cap",
    "∀": r"\forall",
    "∃": r"\exists",
    "∅": r"\emptyset",
}

# n-ary operator mapping
NARY_OP_MAP: dict[str, str] = {
    "∑": r"\sum",
    "∏": r"\prod",
    "∫": r"\int",
    "⋃": r"\bigcup",
    "⋂": r"\bigcap",
}

# Matrix environment to bracket mapping
MATRIX_BRACKET_MAP: dict[str, tuple[str, str]] = {
    "bmatrix": ("[", "]"),
    "pmatrix": ("(", ")"),
    "vmatrix": ("|", "|"),
    "Vmatrix": ("‖", "‖"),
    "Bmatrix": ("{", "}"),
    "matrix": ("", ""),
}

# Bracket pair mapping (for recognizing opening-bracket content closing-bracket patterns)
BRACKET_PAIRS: dict[str, str] = {
    "[": "]",
    "(": ")",
    "{": "}",
    "|": "|",
    "‖": "‖",
    "||": "||",
}
