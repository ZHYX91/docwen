"""Smoke tests for the core-resident Word numbering adapter.

Verifies the scheme → Word ``numbering.xml`` translation layer is usable
from core with zero plugin/runtime dependencies. The comprehensive behavioural
suite also imports this core API directly; this file locks the core-level
public surface and a single end-to-end translation.
"""

from __future__ import annotations

import pytest

from docwen_core.text.numbering_word_adapter import (
    STYLE_TO_NUMFMT,
    LevelCompatibility,
    TranslationResult,
    WordNumberingLevel,
    translate_scheme,
)

pytestmark = pytest.mark.unit


def test_public_api_re_exported_from_core() -> None:
    assert callable(translate_scheme)
    assert isinstance(STYLE_TO_NUMFMT, dict)
    assert WordNumberingLevel.__name__ == "WordNumberingLevel"
    assert LevelCompatibility.__name__ == "LevelCompatibility"
    assert TranslationResult.__name__ == "TranslationResult"


def test_translate_scheme_returns_compatibility_result() -> None:
    scheme = {
        "level_1": {"format": "{1.chinese_lower}、"},
        "level_2": {"format": "（{2.chinese_lower}）"},
    }
    result = translate_scheme(scheme)
    assert isinstance(result, TranslationResult)
    assert result.verdict in {"full", "approximate", "unsupported"}
