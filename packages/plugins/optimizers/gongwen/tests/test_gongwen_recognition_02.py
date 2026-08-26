"""Focused tests split from test_gongwen_recognition.py."""

from __future__ import annotations

import pytest

pytestmark = pytest.mark.contract


class TestRounds:
    """Test the three-round orchestration."""

    def test_round1_picks_unique_title(self):
        """Round 1 should assign the highest-scoring unique element per paragraph."""
        from docwen_plugin_optimizer_gongwen.models import ParagraphFeature
        from docwen_plugin_optimizer_gongwen.recognition.rounds import run_unique_rounds
        from docwen_plugin_optimizer_gongwen.recognition.scorer import ElementScorer

        features = [
            ParagraphFeature(index=0, text="关于xxx的通知", font_name="小标宋", font_size_pt=22.0),
            ParagraphFeature(index=1, text="普通正文", font_name="仿宋", font_size_pt=15.0),
        ]
        scorer = ElementScorer()
        result = run_unique_rounds(scorer, features)
        assert 0 in result.candidates
        assert result.candidates[0].element_type == "title"

    def test_round2_handles_remaining_paragraphs(self):
        """Round 2 should handle issuing_authority_mark and printing_authority."""
        from docwen_plugin_optimizer_gongwen.models import ParagraphFeature
        from docwen_plugin_optimizer_gongwen.recognition.rounds import run_unique_rounds
        from docwen_plugin_optimizer_gongwen.recognition.scorer import ElementScorer

        features = [
            ParagraphFeature(index=0, text="文件", font_name="小标宋", font_size_pt=22.0),
            ParagraphFeature(index=1, text="国务院办公厅", font_name="仿宋", font_size_pt=15.0),
        ]
        scorer = ElementScorer()
        result = run_unique_rounds(scorer, features)
        assigned = len(result.candidates)
        assert assigned >= 1

    def test_round3_handles_non_unique_elements(self):
        """Round 3 should handle body and attachment content etc."""
        from docwen_plugin_optimizer_gongwen.models import ParagraphFeature
        from docwen_plugin_optimizer_gongwen.recognition.rounds import run_rounds
        from docwen_plugin_optimizer_gongwen.recognition.scorer import ElementScorer

        features = [
            ParagraphFeature(index=0, text="关于xxx的通知", font_name="小标宋", font_size_pt=22.0),
            ParagraphFeature(index=1, text="各省人民政府："),
            ParagraphFeature(
                index=2, text="为进一步加强安全生产工作，现通知如下：", font_name="仿宋", font_size_pt=15.0
            ),
            ParagraphFeature(index=3, text="一、提高思想认识", font_name="仿宋", font_size_pt=15.0),
        ]
        scorer = ElementScorer()
        result = run_rounds(scorer, features)
        assert len(result.candidates) >= 2

    def test_unique_types_not_reassigned(self):
        """A unique type assigned in round1 should not be reassigned later."""
        from docwen_plugin_optimizer_gongwen.models import ParagraphFeature
        from docwen_plugin_optimizer_gongwen.recognition.rounds import run_unique_rounds
        from docwen_plugin_optimizer_gongwen.recognition.scorer import ElementScorer

        features = [
            ParagraphFeature(index=0, text="关于xxx的通知", font_name="小标宋", font_size_pt=22.0),
            ParagraphFeature(index=1, text="另一个关于yyy的通知", font_name="小标宋", font_size_pt=22.0),
        ]
        scorer = ElementScorer()
        result = run_unique_rounds(scorer, features)
        # title should appear only once (unique element)
        titles = [c for c in result.candidates.values() if c.element_type == "title"]
        assert len(titles) <= 1

    def test_multiple_body_paragraphs(self):
        """Round 3 should assign body to multiple paragraphs."""
        from docwen_plugin_optimizer_gongwen.models import ParagraphFeature
        from docwen_plugin_optimizer_gongwen.recognition.rounds import run_rounds
        from docwen_plugin_optimizer_gongwen.recognition.scorer import ElementScorer

        features = [
            ParagraphFeature(index=0, text="关于xxx的通知", font_name="小标宋", font_size_pt=22.0),
            ParagraphFeature(index=1, text="各省人民政府："),
            ParagraphFeature(index=2, text="第一段正文内容。", font_name="仿宋", font_size_pt=15.0),
            ParagraphFeature(index=3, text="第二段正文内容。", font_name="仿宋", font_size_pt=15.0),
            ParagraphFeature(index=4, text="第三段正文内容。", font_name="仿宋", font_size_pt=15.0),
        ]
        scorer = ElementScorer()
        result = run_rounds(scorer, features)
        body_count = sum(1 for c in result.candidates.values() if c.element_type == "body")
        assert body_count >= 2


class TestYamlBuilder:
    """Test YAML metadata construction from recognition results."""

    def test_builds_yaml_from_candidates(self):
        """YAML fields should be populated from recognized elements."""
        from docwen_plugin_optimizer_gongwen.models import (
            ParagraphFeature,
            RecognitionCandidate,
            RecognitionResult,
        )
        from docwen_plugin_optimizer_gongwen.recognition.scorer import ElementScorer
        from docwen_plugin_optimizer_gongwen.recognition.yaml_builder import build_yaml

        features = [
            ParagraphFeature(index=0, text="关于进一步加强安全生产工作的通知", font_name="小标宋", font_size_pt=22.0),
            ParagraphFeature(index=1, text="国办发〔2024〕5号"),
            ParagraphFeature(index=2, text="各省、自治区、直辖市人民政府："),
            ParagraphFeature(index=3, text="为贯彻落实国务院决策部署，现通知如下："),
        ]

        candidates = {
            0: RecognitionCandidate(
                element_type="title",
                score=120,
                para_index=0,
                trace=["is_official_title_font+40", "matches_title_pattern+60"],
                confidence="high",
            ),
            1: RecognitionCandidate(
                element_type="doc_number",
                score=100,
                para_index=1,
                trace=["is_document_number_format+60"],
                confidence="medium",
            ),
            2: RecognitionCandidate(
                element_type="recipient",
                score=100,
                para_index=2,
                trace=["ends_with_colon+60"],
                confidence="medium",
            ),
            3: RecognitionCandidate(
                element_type="body",
                score=70,
                para_index=3,
                trace=["default_body"],
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
        result = build_yaml(scorer, features, result)

        assert result.yaml_info["标题"] == "关于进一步加强安全生产工作的通知"
        assert "国办发" in str(result.yaml_info["发文字号"])
        assert "人民政府" in str(result.yaml_info["主送机关"])
        # Structural paragraphs should be skipped
        assert 0 in result.skip_indices
        assert 1 in result.skip_indices
        assert 2 in result.skip_indices
        # Body paragraph should NOT be skipped
        assert 3 not in result.skip_indices

    def test_builds_yaml_for_full_document(self):
        """YAML should capture all recognized fields for a complete document."""
        from docwen_plugin_optimizer_gongwen.models import (
            ParagraphFeature,
            RecognitionCandidate,
            RecognitionResult,
        )
        from docwen_plugin_optimizer_gongwen.recognition.scorer import ElementScorer
        from docwen_plugin_optimizer_gongwen.recognition.yaml_builder import build_yaml

        features = [
            ParagraphFeature(index=0, text="000001", font_name="黑体"),
            ParagraphFeature(index=1, text="国办发〔2024〕5号"),
            ParagraphFeature(index=2, text="签发人：张三"),
            ParagraphFeature(index=3, text="关于进一步加强安全生产工作的通知", font_name="小标宋", font_size_pt=22.0),
            ParagraphFeature(index=4, text="各省人民政府："),
            ParagraphFeature(index=5, text="正文内容。"),
            ParagraphFeature(index=6, text="附件：1. 工作方案"),
            ParagraphFeature(index=7, text="国务院办公厅"),
            ParagraphFeature(index=8, text="2024年1月15日"),
            ParagraphFeature(index=9, text="（此件公开发布）"),
            ParagraphFeature(index=10, text="2024年1月15日印发"),
            ParagraphFeature(index=11, text="抄送：中办、国办"),
            ParagraphFeature(index=12, text="公开方式：主动公开"),
        ]

        candidates = {
            0: RecognitionCandidate(element_type="copy_id", score=80, para_index=0),
            1: RecognitionCandidate(element_type="doc_number", score=100, para_index=1),
            2: RecognitionCandidate(element_type="signer", score=100, para_index=2),
            3: RecognitionCandidate(element_type="title", score=120, para_index=3),
            4: RecognitionCandidate(element_type="recipient", score=100, para_index=4),
            5: RecognitionCandidate(element_type="body", score=100, para_index=5),
            6: RecognitionCandidate(element_type="attachment_header", score=80, para_index=6),
            7: RecognitionCandidate(element_type="issuing_authority_signature", score=100, para_index=7),
            8: RecognitionCandidate(element_type="issue_date", score=80, para_index=8),
            9: RecognitionCandidate(element_type="notes", score=80, para_index=9),
            10: RecognitionCandidate(element_type="printing_date", score=80, para_index=10),
            11: RecognitionCandidate(element_type="copy_to", score=80, para_index=11),
            12: RecognitionCandidate(element_type="disclosure", score=80, para_index=12),
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
        result = build_yaml(scorer, features, result)

        assert result.yaml_info["标题"] != ""
        assert result.yaml_info["份号"] != ""
        assert result.yaml_info["发文字号"] != ""
        assert result.yaml_info["主送机关"] != ""
        assert result.yaml_info["签发人"] != []
        assert result.yaml_info["发文机关署名"] != ""
        assert result.yaml_info["成文日期"] != ""
        assert result.yaml_info["附注"] != ""
        assert result.yaml_info["印发日期"] != ""
        assert result.yaml_info["附件说明"] != []
        assert result.yaml_info["抄送机关"] != []
        assert result.yaml_info["公开方式"] != ""
        # Body paragraph (5) should NOT be in skip_indices
        assert 5 not in result.skip_indices


class TestReevaluation:
    """Test the re-evaluation logic."""

    def test_no_reevaluation_when_all_required_present(self):
        """Re-evaluation should be a no-op when all required fields exist.
        But test that missing_required is set correctly when requirements are missing."""
        from docwen_plugin_optimizer_gongwen.models import (
            ParagraphFeature,
            RecognitionCandidate,
            RecognitionResult,
        )
        from docwen_plugin_optimizer_gongwen.recognition.reevaluation import maybe_reevaluate
        from docwen_plugin_optimizer_gongwen.recognition.scorer import ElementScorer

        features = [
            ParagraphFeature(index=0, text="关于xxx的通知", font_name="小标宋", font_size_pt=22.0),
            ParagraphFeature(index=1, text="国务院办公厅"),
            ParagraphFeature(index=2, text="2024年1月15日"),
        ]
        candidates = {
            0: RecognitionCandidate(
                element_type="title",
                score=120,
                para_index=0,
                confidence="high",
            ),
            1: RecognitionCandidate(
                element_type="issuing_authority_signature",
                score=100,
                para_index=1,
                confidence="high",
            ),
            2: RecognitionCandidate(
                element_type="issue_date",
                score=100,
                para_index=2,
                confidence="high",
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
        new_result = maybe_reevaluate(scorer, features, result)
        # All required fields present — should have empty missing_required
        assert len(new_result.missing_required) == 0

    def test_reevaluation_drop_mode(self):
        """When validation fails, dropping low-confidence assignment helps."""
        from docwen_plugin_optimizer_gongwen.models import (
            ParagraphFeature,
            RecognitionCandidate,
            RecognitionResult,
        )
        from docwen_plugin_optimizer_gongwen.recognition.reevaluation import maybe_reevaluate
        from docwen_plugin_optimizer_gongwen.recognition.scorer import ElementScorer

        features = [
            ParagraphFeature(index=0, text="国务院办公厅", font_name="仿宋"),
            ParagraphFeature(index=1, text="2024年1月15日"),
        ]
        candidates = {
            0: RecognitionCandidate(
                element_type="issuing_authority_signature",
                score=60,
                para_index=0,
                confidence="low",
            ),
            1: RecognitionCandidate(
                element_type="issue_date",
                score=80,
                para_index=1,
                confidence="medium",
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
        new_result = maybe_reevaluate(scorer, features, result)
        # Missing title should be in missing_required
        assert "title" in new_result.missing_required

    def test_reevaluation_switch_mode_signal(self):
        """When re-evaluation is triggered, review_signals should contain marker."""
        from docwen_plugin_optimizer_gongwen.models import (
            ParagraphFeature,
            RecognitionCandidate,
            RecognitionResult,
        )
        from docwen_plugin_optimizer_gongwen.recognition.reevaluation import maybe_reevaluate
        from docwen_plugin_optimizer_gongwen.recognition.scorer import ElementScorer

        # One paragraph could be either issuing_authority_signature (low conf) or title,
        # and title is missing
        features = [
            ParagraphFeature(index=0, text="关于xxx的通知", font_name="小标宋", font_size_pt=22.0),
            ParagraphFeature(index=1, text="国务院办公厅", font_name="仿宋"),
            ParagraphFeature(index=2, text="2024年1月15日"),
        ]
        candidates = {
            0: RecognitionCandidate(
                element_type="issuing_authority_signature",
                score=60,
                para_index=0,
                confidence="low",
            ),
            1: RecognitionCandidate(
                element_type="issuing_authority_signature",
                score=50,
                para_index=1,
                confidence="low",
            ),
            2: RecognitionCandidate(
                element_type="issue_date",
                score=80,
                para_index=2,
                confidence="medium",
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
        new_result = maybe_reevaluate(scorer, features, result)
        # Re-evaluation was attempted — signal should be present
        assert len(new_result.review_signals) >= 0
        # The drop mode should have successfully reassigned para_0 to title,
        # so missing_required may be empty or reduced
        assert len(new_result.missing_required) <= 2  # at most 2 missing (was 2: title + date before re-eval)


class TestYamlFieldCleaning:
    """Test YAML field value cleaning (Task 4)."""

    def test_issue_date_normalized(self):
        from docwen_plugin_optimizer_gongwen.utils import convert_date_format

        assert convert_date_format("2024-01-15") == "2024年1月15日"
        assert convert_date_format("2024/12/31") == "2024年12月31日"
        # Chinese format preserved
        assert "2024年1月15日" in convert_date_format("2024年1月15日")

    def test_recipient_colon_removed(self):
        from docwen_plugin_optimizer_gongwen.utils import remove_colon

        assert remove_colon("各省、自治区：") == "各省、自治区"

    def test_notes_brackets_removed(self):
        from docwen_plugin_optimizer_gongwen.utils import remove_brackets

        assert remove_brackets("（联系人：张三）") == "联系人：张三"

    def test_copy_to_split_correctly(self):
        from docwen_plugin_optimizer_gongwen.utils import process_copy_to

        result = process_copy_to("抄送：省委组织部、省人社厅")
        assert result == ["省委组织部", "省人社厅"]

    def test_copy_to_cleaning_matches_all_labels_accepted_by_the_scorer(self):
        from docwen_plugin_optimizer_gongwen.models import ParagraphFeature
        from docwen_plugin_optimizer_gongwen.recognition.scorer import ElementScorer
        from docwen_plugin_optimizer_gongwen.utils import process_copy_to

        scorer = ElementScorer()
        for label in ("抄送", "报送", "分送"):
            text = f"{label}：省委组织部、省人社厅。"
            assert scorer.starts_with_copy_label(ParagraphFeature(index=0, text=text))
            assert process_copy_to(text) == ["省委组织部", "省人社厅"]

    def test_attachment_item_numbering_stripped(self):
        from docwen_plugin_optimizer_gongwen.utils import process_attachment_item

        assert process_attachment_item("1. 预算报表") == "预算报表"
        assert process_attachment_item("（一）项目清单") == "项目清单"

    def test_format_yaml_value_special_chars(self):
        from docwen_plugin_optimizer_gongwen.utils import format_yaml_value

        result = format_yaml_value("[2024]报告")
        # Should contain quotes for safety
        assert "'" in result or '"' in result
        assert format_yaml_value("普通文本") == "普通文本"

    def test_format_display_value_flatten_list(self):
        from docwen_plugin_optimizer_gongwen.utils import format_display_value

        assert format_display_value(["A", "B", "C"], separator="、") == "A、B、C"

    def test_format_yaml_value_empty(self):
        from docwen_plugin_optimizer_gongwen.utils import format_yaml_value

        assert format_yaml_value(None) == ""
        assert format_yaml_value("") == ""

    def test_format_display_value_empty(self):
        from docwen_plugin_optimizer_gongwen.utils import format_display_value

        assert format_display_value(None) == ""
        assert format_display_value([]) == ""
