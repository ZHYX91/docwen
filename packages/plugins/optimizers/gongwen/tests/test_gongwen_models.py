"""Tests for gongwen data models."""

from __future__ import annotations

import pytest

pytestmark = pytest.mark.unit


def test_paragraph_feature_holds_font_info():
    from docwen_plugin_optimizer_gongwen.models import ParagraphFeature

    pf = ParagraphFeature(
        index=0,
        text="关于xxx的通知",
        font_name="小标宋",
        font_size_pt=22.0,
        style_name="Heading 1",
        outline_level=0,
        alignment="CENTER",
        is_first_in_section=False,
    )
    assert pf.index == 0
    assert pf.font_size_pt == 22.0
    assert pf.style_name == "Heading 1"


def test_recognition_candidate_sorts_by_score():
    from docwen_plugin_optimizer_gongwen.models import RecognitionCandidate

    a = RecognitionCandidate(element_type="标题", score=120, para_index=0)
    b = RecognitionCandidate(element_type="发文机关标志", score=80, para_index=0)
    assert a.score > b.score


def test_gongwen_metadata_defaults_empty():
    from docwen_plugin_optimizer_gongwen.models import GongwenMetadata

    m = GongwenMetadata.default()
    assert m.title == ""
    assert m.doc_number == ""
    assert m.signer == []


def test_gongwen_metadata_has_18_fields():
    from docwen_plugin_optimizer_gongwen.models import GongwenMetadata

    m = GongwenMetadata.default()
    d = m.to_dict()
    assert len(d) == 18
