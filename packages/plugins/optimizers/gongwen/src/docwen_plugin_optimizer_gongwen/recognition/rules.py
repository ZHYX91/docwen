"""Scoring rules table for gongwen element recognition.

Each element type has a list of ScoringRule(condition, score) pairs.
Conditions are method names on ElementScorer.
Positive scores = evidence for, negative scores = evidence against.
"""

from __future__ import annotations

from docwen_plugin_optimizer_gongwen.models import ScoringRule

# ── Scoring Rules Table ──────────────────────────────────────────────
# Organized by element type.  Each rule references a method on ElementScorer
# that accepts (self, pf: ParagraphFeature) -> bool.

SCORING_RULES: dict[str, list[ScoringRule]] = {
    # ── 份号+发文字号 组合 ──
    "combined_id": [
        ScoringRule("has_combined_ids", 100),
    ],
    # ── 份号 (纯数字序列) ──
    "copy_id": [
        ScoringRule("is_numeric_sequence", 60),
        ScoringRule("is_after_last_unique_element", 40),
    ],
    # ── 密级和保密期限 ──
    "security": [
        ScoringRule("starts_with_security_keyword", 60),
        ScoringRule("is_after_last_unique_element", 40),
    ],
    # ── 紧急程度 ──
    "urgency": [
        ScoringRule("starts_with_urgency_keyword", 60),
        ScoringRule("is_after_last_unique_element", 40),
    ],
    # ── 发文机关标志 ──
    "issuing_authority_mark": [
        ScoringRule("ends_with_authority_suffix", 50),
        ScoringRule("is_after_last_unique_element", 30),
        ScoringRule("within_first_3_paragraphs", 30),
    ],
    # ── 发文字号 ──
    "doc_number": [
        ScoringRule("is_document_number_format", 60),
        ScoringRule("is_after_last_unique_element", 40),
    ],
    # ── 发文字号+签发人 组合行 ──
    "combined_doc_number_signer": [
        ScoringRule("has_doc_number_and_signer", 80),
        ScoringRule("follows_issuing_authority_mark", 20),
    ],
    # ── 签发人 ──
    "signer": [
        ScoringRule("starts_with_signer_label", 60),
        ScoringRule("follows_issuing_authority_mark", 20),
        ScoringRule("is_after_last_unique_element", 20),
    ],
    # ── 标题 ──
    "title": [
        ScoringRule("is_official_title_font", 40),
        ScoringRule("is_official_title_size", 40),
        ScoringRule("is_after_last_unique_element", 20),
        ScoringRule("matches_title_pattern", 60),
    ],
    # ── 主送机关 ──
    "recipient": [
        ScoringRule("is_official_title_font", -10),
        ScoringRule("is_official_title_size", -10),
        ScoringRule("ends_with_colon", 60),
        ScoringRule("contains_recipient_chars", 20),
        ScoringRule("is_after_last_unique_element", 20),
    ],
    # ── 附件说明 (第1个附件) ──
    "attachment_header": [
        ScoringRule("is_official_title_font", -10),
        ScoringRule("is_official_title_size", -10),
        ScoringRule("starts_with_attachment_label", 60),
        ScoringRule("is_after_last_unique_element", 20),
        ScoringRule("is_first_attachment", 20),
        ScoringRule("is_too_close_to_recipient_or_title", -20),
    ],
    # ── 发文机关署名 ──
    "issuing_authority_signature": [
        ScoringRule("is_official_title_font", -10),
        ScoringRule("is_official_title_size", -10),
        ScoringRule("ends_with_authority_suffix", 60),
        ScoringRule("follows_attachment", 20),
        ScoringRule("is_after_last_unique_element", 20),
        ScoringRule("is_too_close_to_recipient_or_title", -20),
        ScoringRule("within_first_5_paragraphs", -20),
        ScoringRule("is_table_cell", -60),
        # A copy-to edition line often ends in an authority suffix such as
        # “办公厅”; the explicit label is stronger evidence than that suffix.
        ScoringRule("starts_with_copy_label", -100),
    ],
    # ── 成文日期 ──
    "issue_date": [
        ScoringRule("is_official_title_font", -10),
        ScoringRule("is_official_title_size", -10),
        ScoringRule("is_standalone_date", 60),
        ScoringRule("follows_issuing_authority_signature", 20),
        ScoringRule("is_after_last_unique_element", 20),
        ScoringRule("is_too_close_to_recipient_or_title", -20),
    ],
    # ── 附注 ──
    "notes": [
        ScoringRule("is_official_title_font", -10),
        ScoringRule("is_official_title_size", -10),
        ScoringRule("is_wrapped_in_brackets", 50),
        ScoringRule("follows_issue_date", 30),
        ScoringRule("is_after_last_unique_element", 20),
        ScoringRule("is_too_close_to_recipient_or_title", -20),
    ],
    # ── 公开方式 ──
    "disclosure": [
        ScoringRule("is_official_title_font", -10),
        ScoringRule("is_official_title_size", -10),
        ScoringRule("starts_with_disclosure_label", 60),
        ScoringRule("is_after_last_unique_element", 40),
        ScoringRule("is_too_close_to_recipient_or_title", -20),
    ],
    # ── 抄送机关 ──
    "copy_to": [
        ScoringRule("is_official_title_font", -10),
        ScoringRule("is_official_title_size", -10),
        ScoringRule("starts_with_copy_label", 60),
        ScoringRule("is_after_last_unique_element", 40),
        ScoringRule("is_too_close_to_recipient_or_title", -20),
    ],
    # ── 印发机关 ──
    "printing_authority": [
        ScoringRule("follows_printing_date_reverse", 40),
        ScoringRule("precedes_printing_date_directly", 40),
        ScoringRule("ends_with_authority_suffix", 40),
    ],
    # ── 印发日期 ──
    "printing_date": [
        ScoringRule("is_printing_date_format", 80),
    ],
    # ── 标题后续部分 ──
    "title_following": [
        ScoringRule("is_official_title_font", 40),
        ScoringRule("is_official_title_size", 40),
        ScoringRule("follows_title_directly", 40),
    ],
    # ── 副标题 ──
    "subtitle": [
        ScoringRule("starts_with_subtitle_dash", 60),
        ScoringRule("follows_title_or_title_following", 40),
        ScoringRule("is_official_title_font", -20),
        ScoringRule("is_official_title_size", -20),
    ],
    # ── 副标题后续部分 ──
    "subtitle_following": [
        ScoringRule("follows_subtitle_or_subtitle_following", 60),
        ScoringRule("not_starts_with_dash", 20),
        ScoringRule("is_official_title_font", -20),
        ScoringRule("is_official_title_size", -20),
    ],
    # ── 正文区域 ──
    "body": [
        ScoringRule("is_body_position", 100),
    ],
    # ── 后续附件 ──
    "attachment_following": [
        ScoringRule("is_following_attachment", 100),
    ],
    # ── 签发人后续部分 ──
    "signer_following": [
        ScoringRule("follows_signer_or_signer_following", 60),
        ScoringRule("is_person_name_format", 40),
    ],
    # ── 发文字号+签发人后续部分 ──
    "combined_doc_number_signer_following": [
        ScoringRule("has_doc_number_and_name", 60),
        ScoringRule("follows_combined_doc_number_signer_directly", 40),
    ],
    # ── 附件内容 ──
    "attachment_content": [
        ScoringRule("is_after_last_known_element", 100),
    ],
}
