"""LaTeX to MathML through the shared, complete TeX parser."""

from __future__ import annotations

import logging
import re

from docwen_core.formula.constants import LATEX2MATHML_AVAILABLE, _latex_to_mathml_convert

logger = logging.getLogger(__name__)
_UNMODELED_ROMAN_COMMAND = re.compile(r"(?<!\\)(?:\\\\)*\\(?:mathrm|rm)(?![A-Za-z])")


def latex_to_mathml(latex_str: str) -> str | None:
    """Convert LaTeX, including matrix cells, without ad-hoc token splitting.

    A second handwritten matrix parser used to turn compound cells into
    identifiers and retry failed input as a superficially successful matrix.
    All environments now have the same parser and failure contract.
    """
    if not LATEX2MATHML_AVAILABLE:
        logger.error("latex2mathml library not installed")
        return None
    if not latex_str.strip():
        return None
    # latex2mathml's font-token normalization currently erases the normal
    # variant of ASCII identifiers before emitting MathML. Refuse that known
    # lossy path instead of returning a visually different variable as success.
    if _UNMODELED_ROMAN_COMMAND.search(latex_str):
        logger.warning("Unsupported upright identifier command; preserve the complete source formula")
        return None
    try:
        return _latex_to_mathml_convert(latex_str.strip())
    except Exception as exc:
        logger.error("LaTeX to MathML failed: %s", type(exc).__name__)
        return None
