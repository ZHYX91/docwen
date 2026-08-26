import pytest

from docwen_core.docx_parsing.textbox_extraction import _dedupe_key

pytestmark = pytest.mark.unit


def test_textbox_dedupe_key_uses_text_anchor_and_source():
    assert _dedupe_key("  同一文本  ", 3, "textbox") == ("同一文本", 3, "textbox")
    assert _dedupe_key("同一文本", None, "header_textbox") == (
        "同一文本",
        -1,
        "header_textbox",
    )
