"""Focused tests split from test_gongwen_recognition.py."""

from __future__ import annotations

import pytest

pytestmark = pytest.mark.contract


class TestStructuredSignals:
    """Test structured diagnostic signals (Task 5)."""

    def test_low_confidence_pass_recorded(self):
        """Low confidence assignments generate low_confidence_passes entries."""
        from docwen_plugin_optimizer_gongwen.models import (
            ParagraphFeature,
            RecognitionCandidate,
            RecognitionResult,
        )
        from docwen_plugin_optimizer_gongwen.recognition.signals import collect_structured_signals

        features = [
            ParagraphFeature(index=0, text="普通正文段落"),
        ]
        candidates = {
            0: RecognitionCandidate(
                element_type="body",
                score=60,
                para_index=0,
                confidence="low",
            ),
        }
        result = RecognitionResult(
            candidates=candidates,
            yaml_info={},
            skip_indices=[],
            review_signals=[],
            missing_required=[],
            validation_finding_count=0,
        )
        signals = collect_structured_signals(result, features)
        assert signals["low_confidence_passes"] == [
            {"paragraph_index": 0, "element": "body", "round": 3, "score": 60, "threshold": 60}
        ]

    def test_missing_required_generates_warning(self):
        """Missing required fields produce gongwen_warnings entries."""
        from docwen_plugin_optimizer_gongwen.models import (
            ParagraphFeature,
            RecognitionCandidate,
            RecognitionResult,
        )
        from docwen_plugin_optimizer_gongwen.recognition.signals import collect_structured_signals

        features = [
            ParagraphFeature(index=0, text="text"),
        ]
        candidates = {
            0: RecognitionCandidate(
                element_type="body",
                score=60,
                para_index=0,
                confidence="low",
            ),
        }
        result = RecognitionResult(
            candidates=candidates,
            yaml_info={},
            skip_indices=[],
            review_signals=[],
            missing_required=["title"],
            validation_finding_count=1,
        )
        signals = collect_structured_signals(result, features)
        assert any(w["code"] == "GW001" for w in signals["gongwen_warnings"])

    def test_structured_signals_includes_summary(self):
        """Structured signals should include a recognition_summary."""
        from docwen_plugin_optimizer_gongwen.models import (
            ParagraphFeature,
            RecognitionCandidate,
            RecognitionResult,
        )
        from docwen_plugin_optimizer_gongwen.recognition.signals import collect_structured_signals

        features = [
            ParagraphFeature(index=0, text="text"),
        ]
        candidates = {
            0: RecognitionCandidate(
                element_type="body",
                score=60,
                para_index=0,
                confidence="low",
            ),
        }
        result = RecognitionResult(
            candidates=candidates,
            yaml_info={},
            skip_indices=[],
            review_signals=[],
            missing_required=[],
            validation_finding_count=0,
        )
        signals = collect_structured_signals(result, features)
        assert "recognition_summary" in signals
        assert "status" in signals["recognition_summary"]
        assert "recognized_paragraph_count" in signals["recognition_summary"]

    def test_pipeline_returns_structured_signals(self):
        """Pipeline return dict should include recognition_review_signals."""
        from docx import Document

        from docwen_plugin_optimizer_gongwen.pipeline import convert_docx_to_md_gongwen

        doc = Document()
        doc.add_paragraph("关于xxx的通知")
        result = convert_docx_to_md_gongwen(doc, "test.docx", {})
        assert "recognition_review_signals" in result["metadata"]
        assert "review_signals" not in result["metadata"]
        signals = result["metadata"]["recognition_review_signals"]
        assert "close_unique_matches" in signals
        assert "low_confidence_passes" in signals
        assert "gongwen_warnings" in signals
        assert "recognition_summary" in signals

    def test_second_round_close_match_uses_available_alternatives_at_assignment(self, monkeypatch):
        from docwen_plugin_optimizer_gongwen.models import ParagraphFeature, ScoringRule
        from docwen_plugin_optimizer_gongwen.recognition.rounds import run_rounds
        from docwen_plugin_optimizer_gongwen.recognition.rules import SCORING_RULES
        from docwen_plugin_optimizer_gongwen.recognition.scorer import ElementScorer
        from docwen_plugin_optimizer_gongwen.recognition.signals import collect_structured_signals

        monkeypatch.setitem(SCORING_RULES, "issuing_authority_mark", [ScoringRule("is_after_last_unique_element", 95)])
        monkeypatch.setitem(SCORING_RULES, "printing_authority", [ScoringRule("is_after_last_unique_element", 85)])
        features = [ParagraphFeature(index=0, text="示例机关")]
        result = run_rounds(ElementScorer(), features)
        signals = collect_structured_signals(result, features)
        assert signals["close_unique_matches"] == [
            {
                "paragraph_index": 0,
                "element": "issuing_authority_mark",
                "round": 2,
                "score": 95,
                "runner_up_element": "printing_authority",
                "runner_up_score": 85,
                "score_gap": 10,
            }
        ]

    def test_default_body_is_reported_as_round_three_fallback(self):
        from docwen_plugin_optimizer_gongwen.models import ParagraphFeature
        from docwen_plugin_optimizer_gongwen.recognition.rounds import run_rounds
        from docwen_plugin_optimizer_gongwen.recognition.scorer import ElementScorer
        from docwen_plugin_optimizer_gongwen.recognition.signals import collect_structured_signals

        features = [ParagraphFeature(index=0, text="这里是无法分配的正文。")]
        result = run_rounds(ElementScorer(), features)
        signals = collect_structured_signals(result, features)
        assert signals["used_fallback"] == [
            {"paragraph_index": 0, "element": "body", "round": 3, "reason": "default_body"}
        ]
        assert signals["low_confidence_passes"] == []
