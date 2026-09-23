"""Real DOCX parts prove localized wording, author and exact anchors for all checks."""

from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest

from docwen_core.models.proofread import ProofreadRules
from docwen_plugin_proofread.comment_text import SUPPORTED_COMMENT_LOCALES, format_comment, resolve_comment_locale
from docwen_plugin_proofread.text_validator import TextError

from ._proofread_plugin_support import _build_fake_context, _create_test_docx

@pytest.mark.integration
@pytest.mark.parametrize("locale", SUPPORTED_COMMENT_LOCALES)
def test_four_localized_comments_keep_exact_text_anchors(tmp_path, locale):
    from docwen_plugin_proofread.anchor_report import (
        _extract_comments,
        extract_occurrences_from_document_xml,
        read_docx_part,
    )
    from docwen_plugin_proofread.docx_validator import DocxValidator

    source, staging = tmp_path / "source.docx", tmp_path / "staging"
    staging.mkdir()
    _create_test_docx(str(source), ["才料 秘密 １ （"])
    original = source.read_bytes()
    rules = ProofreadRules(
        symbol_pairs=(("（", "）"),),
        symbol_map={"1": ("１",)},
        typos_map={"材料": ("才料",)},
        sensitive_words={"秘密": ()},
    )
    context = _build_fake_context(
        str(source),
        str(staging),
        target_format="docx",
        action_name="validate",
        options={"locale": locale},
        proofread_rules=rules,
    )
    result = DocxValidator().convert(context)
    assert result.success
    output = Path(result.artifacts[0].staging_path)
    comments = _extract_comments(output)
    assert len(comments) == 4 and {comment.author for comment in comments} == {"DocWen"}
    texts = [comment.text for comment in comments]
    assert sum(" → " in text for text in texts) == 2
    assert any("才料 → 材料" in text for text in texts)
    assert any("１ → 1" in text for text in texts)
    assert all(" → " not in text for text in texts if "秘密" in text or "（" in text)
    if locale == "zh_CN":
        assert set(texts) == {
            "错别字：才料 → 材料",
            "符号修正：１ → 1",
            "命中敏感词：‘秘密’。请结合上下文核查。",
            "‘（’未找到匹配符号，请核查。",
        }
    elif locale != "en_US":
        assert all(not text.startswith(("Typo:", "Symbol correction:", "Sensitive word found:")) for text in texts)
    document_xml = read_docx_part(output, "word/document.xml")
    assert document_xml is not None
    ranges, diagnostics = extract_occurrences_from_document_xml(document_xml, 30, False)
    assert sorted((item.start, item.end, item.covered_text) for item in ranges) == [
        (0, 2, "才料"),
        (3, 5, "秘密"),
        (6, 7, "１"),
        (8, 9, "（"),
    ]
    assert not diagnostics.start_without_end_ids and not diagnostics.end_without_start_ids
    assert source.read_bytes() == original


@pytest.mark.unit
@pytest.mark.parametrize(
    "options,snapshot,expected",
    [
        ({"locale": "fr_FR"}, {"gui": {"language": {"locale": "zh_CN"}}}, "fr_FR"),
        ({}, {"gui": {"language": {"locale": "zh_CN"}}}, "zh_CN"),
        ({"locale": "en_US"}, {}, "en_US"),
        ({}, {}, "en_US"),
    ],
)
def test_comment_language_uses_request_then_frozen_configuration(options, snapshot, expected):
    context = SimpleNamespace(request=SimpleNamespace(options=options, config_snapshot=snapshot))
    assert resolve_comment_locale(context) == expected


@pytest.mark.unit
def test_no_replacement_never_uses_category_name_as_replacement():
    error = TextError(0, 1, ")", "Unmatched Symbol", "Unmatched Symbol", "pairing", pairing_reason="crossed")
    assert format_comment(error, "zh_CN") == "‘)’处存在交叉嵌套，请核查符号配对顺序。"
    assert "→" not in format_comment(error, "en_US")
