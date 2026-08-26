"""Tests for yaml_builder title_following/subtitle_following/signer_following concatenation."""

from __future__ import annotations

import pytest

pytestmark = pytest.mark.contract


def _make_feature(index: int, text: str):
    """Create a minimal ParagraphFeature for testing."""
    from docwen_plugin_optimizer_gongwen.models import ParagraphFeature

    return ParagraphFeature(index=index, text=text)


def _make_candidate(element_type: str, score: int, para_index: int):
    """Create a minimal RecognitionCandidate for testing."""
    from docwen_plugin_optimizer_gongwen.models import RecognitionCandidate

    return RecognitionCandidate(
        element_type=element_type,
        score=score,
        para_index=para_index,
    )


def _make_result(
    candidates: dict[int, tuple[str, int]],
    features: list,
) -> tuple:
    """Build a RecognitionResult and call build_yaml.

    Returns the modified yaml_info dict and skip_indices list.
    """
    from docwen_plugin_optimizer_gongwen.models import RecognitionResult
    from docwen_plugin_optimizer_gongwen.recognition.yaml_builder import build_yaml

    cand_dict = {idx: _make_candidate(etype, score, idx) for idx, (etype, score) in candidates.items()}
    result = RecognitionResult(
        candidates=cand_dict,
        yaml_info={
            "aliases": [],
            "标题": "",
            "副标题": "",
            "份号": "",
            "密级和保密期限": "",
            "紧急程度": "",
            "发文字号": "",
            "发文机关标志": "",
            "签发人": [],
            "发文机关署名": "",
            "成文日期": "",
            "印发日期": "",
            "主送机关": "",
            "附注": "",
            "印发机关": "",
            "抄送机关": [],
            "附件说明": [],
            "公开方式": "",
        },
        skip_indices=[],
        review_signals=[],
        missing_required=[],
        validation_finding_count=0,
    )
    build_yaml(scorer=None, features=features, result=result)
    return result.yaml_info, result.skip_indices


class TestTitleFollowing:
    """title_following concatenation behavior."""

    def test_single_segment_title_no_following(self):
        """A single title paragraph is written as-is."""
        features = [_make_feature(0, "关于进一步加强安全生产工作的通知")]
        candidates = {0: ("title", 100)}
        yaml_info, skip = _make_result(candidates, features)
        assert yaml_info["标题"] == "关于进一步加强安全生产工作的通知"
        assert 0 in skip

    def test_two_segment_title_concatenated(self):
        """title + title_following → concatenated into one string."""
        features = [
            _make_feature(0, "关于进一步"),
            _make_feature(1, "加强安全生产工作的通知"),
        ]
        candidates = {0: ("title", 100), 1: ("title_following", 80)}
        yaml_info, skip = _make_result(candidates, features)
        assert yaml_info["标题"] == "关于进一步加强安全生产工作的通知"
        assert 0 in skip
        assert 1 in skip

    def test_three_segment_title_concatenated(self):
        """title + two title_following → all concatenated."""
        features = [
            _make_feature(0, "关于"),
            _make_feature(1, "进一步加强"),
            _make_feature(2, "安全生产工作的通知"),
        ]
        candidates = {
            0: ("title", 100),
            1: ("title_following", 80),
            2: ("title_following", 75),
        }
        yaml_info, _ = _make_result(candidates, features)
        assert yaml_info["标题"] == "关于进一步加强安全生产工作的通知"

    def test_title_following_strips_newlines(self):
        """Newlines and carriage returns are stripped before concatenation."""
        features = [
            _make_feature(0, "关于"),
            _make_feature(1, "xxx\n的通知\r"),
        ]
        candidates = {0: ("title", 100), 1: ("title_following", 80)}
        yaml_info, _ = _make_result(candidates, features)
        assert yaml_info["标题"] == "关于xxx的通知"

    def test_title_following_does_not_pollute_with_whitespace(self):
        """title_following with whitespace-only text does not change title."""
        features = [
            _make_feature(0, "已有标题"),
            _make_feature(1, "  \n  "),
        ]
        candidates = {0: ("title", 100), 1: ("title_following", 80)}
        yaml_info, _ = _make_result(candidates, features)
        assert yaml_info["标题"] == "已有标题"


class TestSubtitleFollowing:
    """subtitle_following concatenation behavior."""

    def test_subtitle_following_concatenated(self):
        """subtitle + subtitle_following → concatenated into one string."""
        features = [
            _make_feature(0, "（讨论稿）"),
            _make_feature(1, "——第二版"),
        ]
        candidates = {0: ("subtitle", 100), 1: ("subtitle_following", 80)}
        yaml_info, skip = _make_result(candidates, features)
        assert yaml_info["副标题"] == "（讨论稿）——第二版"
        assert 0 in skip
        assert 1 in skip

    def test_subtitle_following_strips_newlines(self):
        """Newlines are stripped from subtitle_following."""
        features = [
            _make_feature(0, "副标题"),
            _make_feature(1, "续\n行"),
        ]
        candidates = {0: ("subtitle", 100), 1: ("subtitle_following", 80)}
        yaml_info, _ = _make_result(candidates, features)
        assert yaml_info["副标题"] == "副标题续行"

    def test_subtitle_following_empty_text(self):
        """subtitle_following with empty text does not change subtitle."""
        features = [
            _make_feature(0, "已有副标题"),
            _make_feature(1, "   "),
        ]
        candidates = {0: ("subtitle", 100), 1: ("subtitle_following", 80)}
        yaml_info, _ = _make_result(candidates, features)
        assert yaml_info["副标题"] == "已有副标题"


class TestSignerFollowing:
    """signer_following appending behavior."""

    def test_signer_following_appended(self):
        """signer_following adds new signers to the list."""
        from docwen_plugin_optimizer_gongwen.models import RecognitionResult
        from docwen_plugin_optimizer_gongwen.recognition.yaml_builder import build_yaml

        features = [
            _make_feature(0, "张三"),
            _make_feature(1, "李四、王五"),
        ]
        candidates = {0: ("signer", 100), 1: ("signer_following", 80)}
        cand_dict = {idx: _make_candidate(etype, score, idx) for idx, (etype, score) in candidates.items()}
        yaml_info = {
            "aliases": [],
            "标题": "",
            "副标题": "",
            "份号": "",
            "密级和保密期限": "",
            "紧急程度": "",
            "发文字号": "",
            "发文机关标志": "",
            "签发人": [],
            "发文机关署名": "",
            "成文日期": "",
            "印发日期": "",
            "主送机关": "",
            "附注": "",
            "印发机关": "",
            "抄送机关": [],
            "附件说明": [],
            "公开方式": "",
        }
        result = RecognitionResult(
            candidates=cand_dict,
            yaml_info=yaml_info,
            skip_indices=[],
            review_signals=[],
            missing_required=[],
            validation_finding_count=0,
        )
        build_yaml(scorer=None, features=features, result=result)
        signers = result.yaml_info["签发人"]
        assert "张三" in signers
        assert "李四" in signers
        assert "王五" in signers

    def test_signer_following_dedup(self):
        """Duplicate signers are not added again."""
        from docwen_plugin_optimizer_gongwen.models import RecognitionResult
        from docwen_plugin_optimizer_gongwen.recognition.yaml_builder import build_yaml

        features = [
            _make_feature(0, "张三"),
            _make_feature(1, "张三、李四"),
        ]
        candidates = {0: ("signer", 100), 1: ("signer_following", 80)}
        cand_dict = {idx: _make_candidate(etype, score, idx) for idx, (etype, score) in candidates.items()}
        yaml_info = {
            "aliases": [],
            "标题": "",
            "副标题": "",
            "份号": "",
            "密级和保密期限": "",
            "紧急程度": "",
            "发文字号": "",
            "发文机关标志": "",
            "签发人": [],
            "发文机关署名": "",
            "成文日期": "",
            "印发日期": "",
            "主送机关": "",
            "附注": "",
            "印发机关": "",
            "抄送机关": [],
            "附件说明": [],
            "公开方式": "",
        }
        result = RecognitionResult(
            candidates=cand_dict,
            yaml_info=yaml_info,
            skip_indices=[],
            review_signals=[],
            missing_required=[],
            validation_finding_count=0,
        )
        build_yaml(scorer=None, features=features, result=result)
        assert result.yaml_info["签发人"] == ["张三", "李四"]

    def test_signer_following_empty_text(self):
        """signer_following with empty text does not change signer list."""
        from docwen_plugin_optimizer_gongwen.models import RecognitionResult
        from docwen_plugin_optimizer_gongwen.recognition.yaml_builder import build_yaml

        features = [
            _make_feature(0, "张三"),
            _make_feature(1, "  \n  "),
        ]
        candidates = {0: ("signer", 100), 1: ("signer_following", 80)}
        cand_dict = {idx: _make_candidate(etype, score, idx) for idx, (etype, score) in candidates.items()}
        yaml_info = {
            "aliases": [],
            "标题": "",
            "副标题": "",
            "份号": "",
            "密级和保密期限": "",
            "紧急程度": "",
            "发文字号": "",
            "发文机关标志": "",
            "签发人": ["张三"],
            "发文机关署名": "",
            "成文日期": "",
            "印发日期": "",
            "主送机关": "",
            "附注": "",
            "印发机关": "",
            "抄送机关": [],
            "附件说明": [],
            "公开方式": "",
        }
        result = RecognitionResult(
            candidates=cand_dict,
            yaml_info=yaml_info,
            skip_indices=[],
            review_signals=[],
            missing_required=[],
            validation_finding_count=0,
        )
        build_yaml(scorer=None, features=features, result=result)
        assert result.yaml_info["签发人"] == ["张三"]


class TestExtractSignersFromText:
    """extract_signers_from_text utility."""

    def test_basic_split_comma(self):
        from docwen_plugin_optimizer_gongwen.utils import extract_signers_from_text

        result = extract_signers_from_text("李四、王五")
        assert result == ["李四", "王五"]

    def test_basic_split_space(self):
        from docwen_plugin_optimizer_gongwen.utils import extract_signers_from_text

        result = extract_signers_from_text("李四 王五")
        assert result == ["李四", "王五"]

    def test_minority_name_with_separator_dot(self):
        from docwen_plugin_optimizer_gongwen.utils import extract_signers_from_text

        result = extract_signers_from_text("艾尼·买买提、张三")
        assert result == ["艾尼·买买提", "张三"]

    def test_single_name(self):
        from docwen_plugin_optimizer_gongwen.utils import extract_signers_from_text

        result = extract_signers_from_text("李四")
        assert result == ["李四"]

    def test_invalid_name_filtered(self):
        """Non-Chinese or too-short tokens are filtered out."""
        from docwen_plugin_optimizer_gongwen.utils import extract_signers_from_text

        result = extract_signers_from_text("A、李、张三")
        assert result == ["张三"]

    def test_empty_text(self):
        from docwen_plugin_optimizer_gongwen.utils import extract_signers_from_text

        assert extract_signers_from_text("") == []
        assert extract_signers_from_text("   ") == []


def test_yaml_builder_extracts_composite_and_labelled_fields_exactly():
    features = [
        _make_feature(0, "5国办发〔2024〕5号"),
        _make_feature(1, "国办发〔2024〕5号  签发人：张三、李四"),
        _make_feature(2, "国办发〔2024〕5号  王五"),
        _make_feature(3, "签发人：赵六"),
        _make_feature(4, "公开方式：主动公开"),
        _make_feature(5, "国务院办公厅　2024年1月15日印发"),
        _make_feature(6, "关于进一步"),
        _make_feature(7, "规范公文处理工作的通知"),
    ]
    candidates = {
        0: ("combined_id", 100),
        1: ("combined_doc_number_signer", 100),
        2: ("combined_doc_number_signer_following", 100),
        3: ("signer", 100),
        4: ("disclosure", 100),
        5: ("printing_date", 100),
        6: ("title", 100),
        7: ("title_following", 100),
    }

    yaml_info, skip = _make_result(candidates, features)

    assert yaml_info["份号"] == "5"
    assert yaml_info["发文字号"] == "国办发〔2024〕5号"
    assert yaml_info["签发人"] == ["张三", "李四", "王五", "赵六"]
    assert yaml_info["公开方式"] == "主动公开"
    assert yaml_info["印发机关"] == "国务院办公厅"
    assert yaml_info["印发日期"] == "2024年1月15日"
    assert yaml_info["标题"] == "关于进一步规范公文处理工作的通知"
    assert yaml_info["aliases"] == ["关于进一步规范公文处理工作的通知"]
    assert set(skip) == set(range(8))


def test_attachment_description_is_deduplicated_after_cleaning():
    features = [
        _make_feature(0, "附件：1. 工作方案"),
        _make_feature(1, "1. 工作方案"),
    ]
    candidates = {
        0: ("attachment_header", 100),
        1: ("attachment_following", 100),
    }

    yaml_info, _ = _make_result(candidates, features)

    assert yaml_info["附件说明"] == ["工作方案"]
