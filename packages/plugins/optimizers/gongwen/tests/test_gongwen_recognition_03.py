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
        from docwen_plugin_optimizer_gongwen.recognition.scorer import ElementScorer
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
        scorer = ElementScorer()
        scorer.reset_context(features)
        signals = collect_structured_signals(result, scorer, features)
        assert len(signals["low_confidence_passes"]) >= 1
        for entry in signals["low_confidence_passes"]:
            assert entry["score"] < 100  # low confidence threshold

    def test_missing_required_generates_warning(self):
        """Missing required fields produce gongwen_warnings entries."""
        from docwen_plugin_optimizer_gongwen.models import (
            ParagraphFeature,
            RecognitionCandidate,
            RecognitionResult,
        )
        from docwen_plugin_optimizer_gongwen.recognition.scorer import ElementScorer
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
        scorer = ElementScorer()
        scorer.reset_context(features)
        signals = collect_structured_signals(result, scorer, features)
        assert any(w["code"] == "GW001" for w in signals["gongwen_warnings"])

    def test_structured_signals_includes_summary(self):
        """Structured signals should include a recognition_summary."""
        from docwen_plugin_optimizer_gongwen.models import (
            ParagraphFeature,
            RecognitionCandidate,
            RecognitionResult,
        )
        from docwen_plugin_optimizer_gongwen.recognition.scorer import ElementScorer
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
        scorer = ElementScorer()
        scorer.reset_context(features)
        signals = collect_structured_signals(result, scorer, features)
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
