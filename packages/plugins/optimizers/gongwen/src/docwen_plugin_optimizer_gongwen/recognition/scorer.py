"""Element scoring engine for gongwen recognition.

The scorer is a plain service with an injectable rule table and no mixin
inheritance.

All condition methods accept (self, pf: ParagraphFeature) -> bool and are
referenced by name from SCORING_RULES via getattr.
"""

from __future__ import annotations

import re
from collections.abc import Callable
from typing import TYPE_CHECKING

from docwen_plugin_optimizer_gongwen.constants import (
    NON_UNIQUE_SCORE_THRESHOLD,
    ROUND1_ELEMENTS,
    ROUND2_ELEMENTS,
    ROUND3_ELEMENTS,
    UNIQUE_SCORE_THRESHOLD,
)
from docwen_plugin_optimizer_gongwen.models import ParagraphFeature, RecognitionCandidate
from docwen_plugin_optimizer_gongwen.recognition.rules import SCORING_RULES
from docwen_plugin_optimizer_gongwen.utils import starts_with_copy_to_label

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import ProgressSink

# ── Document number patterns (from old ElementScorer) ─────────────────
_DOCUMENT_NUMBER_PATTERNS = (
    r"[一-龥A-Za-z0-9]+〔\d{4}〕\s*\d*\s*号",
    r"[一-龥A-Za-z0-9]+\[\d{4}\]\s*\d*\s*号",
    r"[一-龥A-Za-z0-9]+\(\d{4}\)\s*\d*\s*号",
    r"[一-龥A-Za-z0-9]+\d{4}-\s*\d*\s*号",
    r"[一-龥A-Za-z0-9]+〔\d{4}〕[\s　]*[\d]*[\s　]*号",
    r"[一-龥A-Za-z0-9]+\[\d{4}\][\s　]*[\d]*[\s　]*号",
)

_DOCUMENT_NUMBER_FULL_PATTERNS = tuple(re.compile(f"^{p}\\s*$") for p in _DOCUMENT_NUMBER_PATTERNS)

# ── Authority suffixes (from old ends_with keyword list) ──────────────
_AUTHORITY_SUFFIXES = [
    "人民政府",
    "委员会",
    "办公室",
    "局",
    "厅",
    "部",
    "院",
    "校",
    "中心",
    "办",
    "组",
    "室",
    "司",
    "署",
    "处",
    "科",
]

# ── Name matching patterns (inlined from old name_utils) ──────────────
_NAME_PART_PATTERN = r"[一-龥]{2,12}"
_NAME_SEPARATOR_CHARS = "·•"
_NAME_PATTERN = rf"{_NAME_PART_PATTERN}(?:[{re.escape(_NAME_SEPARATOR_CHARS)}]{_NAME_PART_PATTERN})*"
_NAME_RE = re.compile(rf"^{_NAME_PATTERN}$")
_NAME_LIST_SPLIT_RE = re.compile(r"[、\s　]+")

# ── Document keywords blacklist (for person name detection) ───────────
_DOC_KEYWORDS = [
    "通知",
    "决定",
    "批复",
    "意见",
    "报告",
    "请示",
    "函",
    "通报",
    "通告",
    "公告",
    "议案",
    "纪要",
    "办法",
    "规定",
    "人民政府",
    "委员会",
    "办公室",
    "关于",
]


def _split_name_list(text: str) -> list[str]:
    """Split name list text by Chinese/whitespace separators."""
    parts = _NAME_LIST_SPLIT_RE.split(text.strip())
    return [p.strip() for p in parts if p.strip()]


def _is_person_name(text: str) -> bool:
    """Check if a single token matches person name pattern."""
    return bool(_NAME_RE.match(text.strip()))


class ElementScorer:
    """Scores paragraphs against gongwen element type rules.

    No mixin inheritance. Plain service with injectable rule table.
    Maintains context state (previous element assignments) for position-based rules.
    """

    def __init__(self, diagnostic_sink: ProgressSink | None = None):
        # Context state
        self._diagnostic_sink = diagnostic_sink
        self._reported_rule_failures: set[tuple[int, str, str, str]] = set()
        self._rule_failures: list[dict[str, int | str]] = []
        self._element_positions: dict[str, int] = {}
        self._last_unique_element: str | None = None
        self._last_attachment_marker_index: int = -1
        self._all_features: list[ParagraphFeature] = []

    # ── Public API ─────────────────────────────────────────────────

    def reset_context(self, features: list[ParagraphFeature]) -> None:
        """Reset all context state for a new document.

        Args:
            features: All paragraph features in document order.
        """
        # Element type names collected from round groups
        all_types = set(ROUND1_ELEMENTS) | set(ROUND2_ELEMENTS) | set(ROUND3_ELEMENTS)
        self._element_positions = dict.fromkeys(all_types, -1)
        self._last_unique_element = None
        self._last_attachment_marker_index = -1
        self._all_features = list(features)
        self._reported_rule_failures.clear()
        self._rule_failures.clear()

    @property
    def rule_failures(self) -> tuple[dict[str, int | str], ...]:
        """Return de-duplicated best-effort rule failures for metrics/tests."""
        return tuple(dict(item) for item in self._rule_failures)

    def update_context(self, element_type: str, para_index: int) -> None:
        """Record that an element was assigned to a paragraph.

        Faithfully mirrors old _ContextCoreMixin.update_context logic:
        - Non-unique elements: always update position (for continuous tracking)
        - Unique elements: only record the first occurrence
        - Track attachment markers separately
        - Track last unique element for ordering checks
        """
        if element_type not in self._element_positions:
            return

        is_non_unique = element_type in ROUND3_ELEMENTS
        is_attachment_marker = element_type in {"attachment_header", "attachment_following"}

        if is_non_unique:
            # Always update non-unique element positions for continuous tracking
            self._element_positions[element_type] = para_index
        elif self._element_positions[element_type] == -1:
            # Only record first occurrence for unique elements
            self._element_positions[element_type] = para_index

        if is_attachment_marker:
            self._last_attachment_marker_index = max(self._last_attachment_marker_index, para_index)

        if element_type not in ROUND3_ELEMENTS:
            self._last_unique_element = element_type

    def score_round(
        self,
        pf: ParagraphFeature,
        round_group: str = "round1",
    ) -> list[RecognitionCandidate]:
        """Score a paragraph against all element types in a round group.

        Args:
            pf: The paragraph feature to score.
            round_group: "round1", "round2", or "round3".

        Returns:
            List of RecognitionCandidate sorted by score descending.
            Only candidates above threshold are included.
        """
        element_sets = {
            "round1": ROUND1_ELEMENTS,
            "round2": ROUND2_ELEMENTS,
            "round3": ROUND3_ELEMENTS,
        }
        threshold = NON_UNIQUE_SCORE_THRESHOLD if round_group == "round3" else UNIQUE_SCORE_THRESHOLD

        element_types = element_sets.get(round_group, ROUND1_ELEMENTS)
        candidates: list[RecognitionCandidate] = []

        for element_type in element_types:
            rules = SCORING_RULES.get(element_type, [])
            if not rules:
                continue

            score = 0
            trace: list[str] = []

            for rule in rules:
                checker: Callable | None = getattr(self, rule.condition, None)
                if checker is None:
                    continue

                try:
                    result = checker(pf)
                    if result:
                        score += rule.score
                        sign = "+" if rule.score >= 0 else ""
                        trace.append(f"{rule.condition}{sign}{rule.score}")
                except Exception as exc:
                    failure_key = (
                        pf.index,
                        element_type,
                        rule.condition,
                        type(exc).__name__,
                    )
                    if failure_key in self._reported_rule_failures:
                        continue
                    self._reported_rule_failures.add(failure_key)
                    self._rule_failures.append(
                        {
                            "paragraph_index": pf.index,
                            "element_type": element_type,
                            "condition": rule.condition,
                            "exception_type": type(exc).__name__,
                        }
                    )
                    if self._diagnostic_sink is not None:
                        self._diagnostic_sink.report_diagnostic(
                            "warning",
                            (
                                f"Skipped Gongwen scoring rule '{rule.condition}' for "
                                f"element '{element_type}' after {type(exc).__name__}; "
                                "recognition may be incomplete."
                            ),
                            code="GONGWEN-SCORING-RULE-SKIPPED",
                            location=f"paragraph {pf.index}",
                        )

            if score >= threshold:
                candidates.append(
                    RecognitionCandidate(
                        element_type=element_type,
                        score=score,
                        para_index=pf.index,
                        trace=trace,
                        confidence=self._score_to_confidence(score),
                    )
                )

        # Python's stable sort preserves the declared round-group priority
        # when candidates have equal scores.
        candidates.sort(key=lambda candidate: candidate.score, reverse=True)
        return candidates

    def _score_to_confidence(self, score: int) -> str:
        """Map numeric score to confidence level."""
        if score >= 120:
            return "high"
        elif score >= 80:
            return "medium"
        return "low"

    # ── Text Pattern Checkers ──────────────────────────────────────

    def has_combined_ids(self, pf: ParagraphFeature) -> bool:
        """Check whether text combines a copy ID and document number."""
        text = pf.text.strip()
        copy_id_patterns = [
            r"\d+[\s　\t]*",
            r"\d+[\s　\t]+\d+[\s　\t]*",
        ]
        for cp in copy_id_patterns:
            for dp in _DOCUMENT_NUMBER_PATTERNS:
                full = re.compile(f"^({cp})({dp})\\s*$")
                m = full.match(text)
                if m:
                    return True
        return False

    def is_numeric_sequence(self, pf: ParagraphFeature) -> bool:
        """Check whether text is a pure numeric copy-ID sequence."""
        return pf.text.strip().isdigit()

    def starts_with_security_keyword(self, pf: ParagraphFeature) -> bool:
        """Check whether text starts with a secrecy-level keyword."""
        t = pf.text.strip()
        return t.startswith(("绝密", "机密", "秘密"))

    def starts_with_urgency_keyword(self, pf: ParagraphFeature) -> bool:
        """Check for supported urgency prefixes, including spaced forms."""
        t = pf.text.strip()
        patterns = [
            r"^特急\b",
            r"^加急\b",
            r"^特提\b",
            r"^平急\b",
            r"^特\s*急",
            r"^加\s*急",
            r"^特\s*提",
            r"^平\s*急",
        ]
        return any(re.search(p, t) for p in patterns)

    def is_document_number_format(self, pf: ParagraphFeature) -> bool:
        """Check whether the whole text matches a document number."""
        text = pf.text.strip()
        return any(pattern.match(text) for pattern in _DOCUMENT_NUMBER_FULL_PATTERNS)

    def has_doc_number_and_signer(self, pf: ParagraphFeature) -> bool:
        """Check for a document number followed by labelled signer names."""
        text = pf.text.strip()
        for dp in _DOCUMENT_NUMBER_PATTERNS:
            pattern = rf"^({dp})[\s　\t]+签发人[：:]\s*{_NAME_PATTERN}(?:[、\s　]+{_NAME_PATTERN})*$"
            if re.match(pattern, text):
                return True
        return False

    def has_doc_number_and_name(self, pf: ParagraphFeature) -> bool:
        """Check for a document number followed by an unlabelled name."""
        text = pf.text.strip()
        for dp in _DOCUMENT_NUMBER_PATTERNS:
            pattern = rf"^({dp})[\s　\t]+{_NAME_PATTERN}$"
            if re.match(pattern, text):
                return True
        return False

    def starts_with_signer_label(self, pf: ParagraphFeature) -> bool:
        """Check whether text starts with a signer label."""
        t = pf.text.strip()
        return t.startswith(("签发人：", "签发人:"))

    def starts_with_attachment_label(self, pf: ParagraphFeature) -> bool:
        """Check whether text starts with a supported attachment marker."""
        t = pf.text.strip()
        return t.startswith(("附件：", "附件:", "附件1：", "附件1:"))

    def starts_with_disclosure_label(self, pf: ParagraphFeature) -> bool:
        """Check whether text starts with a disclosure label."""
        t = pf.text.strip()
        return t.startswith(("公开方式：", "公开方式:", "公开属性：", "公开属性:"))

    def starts_with_copy_label(self, pf: ParagraphFeature) -> bool:
        """Check whether text starts with a copy-distribution label."""
        return starts_with_copy_to_label(pf.text)

    def starts_with_subtitle_dash(self, pf: ParagraphFeature) -> bool:
        """Check whether text starts with a subtitle dash marker."""
        t = pf.text.strip()
        return t.startswith(("——", "──"))

    def not_starts_with_dash(self, pf: ParagraphFeature) -> bool:
        """Exclude paragraphs that start a dashed subtitle."""
        t = pf.text.strip()
        return not t.startswith(("——", "──"))

    def matches_title_pattern(self, pf: ParagraphFeature) -> bool:
        """Match a title ending in a supported official-document type."""
        text = pf.text.strip()
        pattern = re.compile(
            r"^(?:[一-龥a-zA-Z0-9]+(?:[\s　]+[一-龥a-zA-Z0-9]+)*[\s　]+)?"
            r"(?:关于)?"
            r"[《]?(.+?)[》]?"
            r"(?:的)?"
            r"[\s　]*"
            r"(?:通知|决定|批复|意见|报告|请示|函|通报|通告|公告|议案|纪要|细则|办法|规定|方案|计划|指示|命令|倡议|倡议书|公示|说明)$"
        )
        return pattern.match(text) is not None

    def is_standalone_date(self, pf: ParagraphFeature) -> bool:
        """Check whether the whole paragraph is a date."""
        text = pf.text.strip()
        date_pattern = r"^\d{4}[年\-\.\/\\]\d{1,2}[月\-\.\/\\]\d{1,2}[日号\-\.\/\\]?$"
        return re.fullmatch(date_pattern, text) is not None

    def is_wrapped_in_brackets(self, pf: ParagraphFeature) -> bool:
        """Check whether text is wrapped as a parenthetical note."""
        text = pf.text.strip()
        if len(text) < 2:
            return False
        left = text[0] in ("(", "（")
        right = text[-1] in (")", "）")
        return left and right

    def is_printing_date_format(self, pf: ParagraphFeature) -> bool:
        """Match a printing date whose explicit suffix distinguishes it from issue dates."""
        text = pf.text.strip()
        patterns = [
            r"^(?:.*?\s*)?\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日\s*印发$",
            r"^(?:.*?\s*)?\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*号\s*印发$",
        ]
        return any(re.match(p, text) for p in patterns)

    def is_person_name_format(self, pf: ParagraphFeature) -> bool:
        """Check whether text represents a list of person names.

        Accepted features:
        - 2-12 consecutive Chinese characters per name
        - May contain 、or spaces separating multiple names
        - Must not contain common document keywords
        - Must not contain punctuation (except 、)
        """
        text = pf.text.strip()
        if text.startswith(("抄送", "报送", "分送", "印发", "公开方式", "公开属性")):
            return False
        if any(kw in text for kw in _DOC_KEYWORDS):
            return False
        if re.search(r'[：:。！？；""' "《》【】()（）]", text):
            return False
        names = _split_name_list(text)
        for name in names:
            if not _is_person_name(name):
                return False
        return len(names) > 0

    def is_numbered_list_item(self, pf: ParagraphFeature) -> bool:
        """Check whether text starts a decimal numbered-list item."""
        return bool(re.match(r"^\d+[\.．]\s*", pf.text.strip()))

    def is_first_attachment(self, pf: ParagraphFeature) -> bool:
        """Check whether no attachment header has been assigned yet."""
        return self._element_positions.get("attachment_header", -1) == -1

    def is_following_attachment(self, pf: ParagraphFeature) -> bool:
        """Check whether a numbered paragraph follows the first attachment.

        Requirements:
        - First attachment has been identified
        - Current paragraph is a numbered list item
        - Current paragraph is after the first attachment
        """
        attach_pos = self._element_positions.get("attachment_header", -1)
        if attach_pos == -1:
            return False
        if not self.is_numbered_list_item(pf):
            return False
        return pf.index > attach_pos

    # ── Text Feature Checkers (ends_with / contains) ───────────────

    def ends_with_authority_suffix(self, pf: ParagraphFeature) -> bool:
        """Check whether text ends with a known organization suffix."""
        t = pf.text.strip()
        return any(t.endswith(s) for s in _AUTHORITY_SUFFIXES)

    def ends_with_colon(self, pf: ParagraphFeature) -> bool:
        """Check whether text ends with a full-width or ASCII colon."""
        t = pf.text.strip()
        return t.endswith(("：", ":"))

    def contains_recipient_chars(self, pf: ParagraphFeature) -> bool:
        """Check for characters commonly present in recipient names."""
        chars = ["各", "委", "局", "办", "厅", "部", "院", "校"]
        return any(c in pf.text for c in chars)

    # ── Font Checkers ──────────────────────────────────────────────

    def is_official_title_font(self, pf: ParagraphFeature) -> bool:
        """Check for supported official-title font families."""
        font = pf.font_name or ""
        return "小标宋" in font or "黑体" in font

    def is_official_title_size(self, pf: ParagraphFeature) -> bool:
        """Check whether font size is within the 22-point title tolerance."""
        if pf.font_size_pt is None:
            return False
        return 21.0 <= pf.font_size_pt <= 23.0

    # ── Position / Follows Checkers ─────────────────────────────────

    def _get_element_pos(self, element_type: str) -> int:
        """Get the recorded paragraph index for an element type."""
        return self._element_positions.get(element_type, -1)

    def _count_nonempty_between(self, start: int, end: int) -> int:
        """Count non-empty paragraphs between start (exclusive) and end (exclusive)."""
        count = 0
        for i in range(start + 1, end):
            if i < len(self._all_features) and self._all_features[i].text.strip():
                count += 1
        return count

    def _follows_element_check(
        self,
        pf: ParagraphFeature,
        preceding_type: str,
        max_nonempty_between: int | None = None,
    ) -> bool:
        """Generic check: does current paragraph follow a preceding element.

        Args:
            pf: Current paragraph feature.
            preceding_type: Element type that should precede this paragraph.
            max_nonempty_between: Max non-empty paragraphs allowed between.
                None = no limit. 0 = must be directly after.
        """
        preceding_pos = self._get_element_pos(preceding_type)
        if preceding_pos == -1:
            return False
        if pf.index <= preceding_pos:
            return False
        if max_nonempty_between is None:
            return True
        return self._count_nonempty_between(preceding_pos, pf.index) <= max_nonempty_between

    def follows_issuing_authority_mark(self, pf: ParagraphFeature) -> bool:
        """Check whether the paragraph follows the issuing-authority mark."""
        return self._follows_element_check(pf, "issuing_authority_mark")

    def follows_title_directly(self, pf: ParagraphFeature) -> bool:
        """Check whether the paragraph directly follows the title."""
        return self._follows_element_check(pf, "title", max_nonempty_between=0)

    def follows_attachment(self, pf: ParagraphFeature) -> bool:
        """Check whether the paragraph follows an attachment header."""
        return self._follows_element_check(pf, "attachment_header")

    def follows_issuing_authority_signature(self, pf: ParagraphFeature) -> bool:
        """Check whether the paragraph follows the issuing-authority signature."""
        return self._follows_element_check(pf, "issuing_authority_signature")

    def follows_issue_date(self, pf: ParagraphFeature) -> bool:
        """Check whether the paragraph follows the issue date."""
        return self._follows_element_check(pf, "issue_date")

    def follows_printing_date_reverse(self, pf: ParagraphFeature) -> bool:
        """Check whether this printing authority precedes the printing date."""
        printing_pos = self._get_element_pos("printing_date")
        if printing_pos == -1:
            return False
        return pf.index < printing_pos

    def precedes_printing_date_directly(self, pf: ParagraphFeature) -> bool:
        """Check that the next non-empty feature is the printing date."""

        printing_pos = self._get_element_pos("printing_date")
        if printing_pos == -1 or pf.index >= printing_pos:
            return False
        return self._count_nonempty_between(pf.index, printing_pos) == 0

    def follows_combined_doc_number_signer_directly(self, pf: ParagraphFeature) -> bool:
        """Check whether the paragraph directly follows a combined number and signer."""
        return self._follows_element_check(pf, "combined_doc_number_signer", max_nonempty_between=0)

    def follows_title_or_title_following(self, pf: ParagraphFeature) -> bool:
        """Check whether the paragraph closely follows the title sequence."""
        title_pos = self._get_element_pos("title")
        title_following_pos = self._get_element_pos("title_following")
        ref_pos = max(title_pos, title_following_pos)
        if ref_pos == -1:
            return False
        if pf.index <= ref_pos:
            return False
        return self._count_nonempty_between(ref_pos, pf.index) <= 2

    def follows_subtitle_or_subtitle_following(self, pf: ParagraphFeature) -> bool:
        """Check whether the paragraph directly follows the subtitle sequence."""
        subtitle_pos = self._get_element_pos("subtitle")
        sub_following_pos = self._get_element_pos("subtitle_following")
        ref_pos = max(subtitle_pos, sub_following_pos)
        if ref_pos == -1:
            return False
        if pf.index <= ref_pos:
            return False
        return self._count_nonempty_between(ref_pos, pf.index) == 0

    def follows_signer_or_signer_following(self, pf: ParagraphFeature) -> bool:
        """Check whether the paragraph directly follows the signer sequence."""
        signer_pos = self._get_element_pos("signer")
        signer_following_pos = self._get_element_pos("signer_following")
        ref_pos = max(signer_pos, signer_following_pos)
        if ref_pos == -1:
            return False
        if pf.index <= ref_pos:
            return False
        return self._count_nonempty_between(ref_pos, pf.index) == 0

    # ── Range Checkers ─────────────────────────────────────────────

    def within_first_3_paragraphs(self, pf: ParagraphFeature) -> bool:
        """Check whether the paragraph is within the first three positions."""
        return pf.index < 3

    def within_first_5_paragraphs(self, pf: ParagraphFeature) -> bool:
        """Check whether the paragraph is within the first five positions."""
        return pf.index < 5

    # ── Ordering / Position Checkers ───────────────────────────────

    def is_after_last_unique_element(self, pf: ParagraphFeature) -> bool:
        """Enforce ordering after the most recently assigned unique element."""
        if self._last_unique_element is None:
            return True
        last_pos = self._element_positions.get(self._last_unique_element, -1)
        if last_pos == -1:
            return True
        return pf.index > last_pos

    def is_body_position(self, pf: ParagraphFeature) -> bool:
        """Check whether a paragraph falls within the document body region.

        The body region is:
        - After recipient (or title if no recipient)
        - Before attachment (if exists)
        - Before issuing_authority_signature (if exists)
        - Before issue_date (if no issuing_authority_signature but issue_date exists)
        """
        start_ref = self._get_element_pos("recipient")
        if start_ref == -1:
            start_ref = self._get_element_pos("title")
        if start_ref == -1:
            return False

        if pf.index <= start_ref:
            return False

        att_pos = self._get_element_pos("attachment_header")
        if att_pos != -1 and pf.index >= att_pos:
            return False

        sig_pos = self._get_element_pos("issuing_authority_signature")
        if sig_pos != -1 and pf.index >= sig_pos:
            return False

        if sig_pos == -1:
            date_pos = self._get_element_pos("issue_date")
            if date_pos != -1 and pf.index >= date_pos:
                return False

        return True

    def is_too_close_to_recipient_or_title(self, pf: ParagraphFeature) -> bool:
        """Check whether the paragraph is fewer than two positions after the heading block."""
        ref_pos = self._get_element_pos("recipient")
        if ref_pos == -1:
            ref_pos = self._get_element_pos("title")
        if ref_pos == -1:
            return False

        distance = pf.index - ref_pos
        return distance < 2

    def is_after_attachment_marker(self, pf: ParagraphFeature) -> bool:
        """Check whether the paragraph appears after an attachment marker."""
        if self._last_attachment_marker_index == -1:
            return False
        return pf.index > self._last_attachment_marker_index

    def is_after_last_known_element(self, pf: ParagraphFeature) -> bool:
        """Check whether attachment body follows its marker and all metadata.

        Merely being after the last recognized element is not sufficient: an
        ordinary trailing paragraph or data-table cell would otherwise become
        an attachment artifact even when the document has no attachment list.
        """

        last_known_index = max(self._element_positions.values(), default=-1)
        return (
            self._last_attachment_marker_index >= 0
            and pf.index > self._last_attachment_marker_index
            and last_known_index >= 0
            and pf.index > last_known_index
        )

    def is_table_cell(self, pf: ParagraphFeature) -> bool:
        """Return True for content flattened from a DOCX table cell."""

        return pf.source == "table" or bool(pf.table_cell_context)
