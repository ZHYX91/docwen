"""Post-recognition validation for gongwen documents."""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from docwen_plugin_optimizer_gongwen.models import RecognitionResult


# Required elements for a valid gongwen document
REQUIRED_ELEMENTS = {"title", "issuing_authority_signature", "issue_date"}

# Structural validation rules (element ordering checks)
STRUCTURAL_RULES = [
    ("title", "body", "title_before_body", "标题应在正文之前"),
    (
        "issuing_authority_signature",
        "issue_date",
        "sig_before_date",
        "发文机关署名应在成文日期之前",
    ),
    ("recipient", "body", "recipient_before_body", "主送机关应在正文之前"),
]


def validate_result(result: RecognitionResult) -> RecognitionResult:
    """Validate recognition result and update findings.

    Checks:
    1. Required elements are present
    2. Structural ordering is correct

    Returns the result with updated validation_finding_count and missing_required.
    """
    found_types = {c.element_type for c in result.candidates.values()}

    # Check required elements
    missing = REQUIRED_ELEMENTS - found_types
    result.missing_required = sorted(missing)

    # Count findings
    finding_count = len(missing)

    # Check structural rules
    for before_type, after_type, _rule_id, message in STRUCTURAL_RULES:
        if before_type in found_types and after_type in found_types:
            before_idx = None
            after_idx = None
            for idx, c in result.candidates.items():
                if c.element_type == before_type:
                    before_idx = idx
                elif c.element_type == after_type:
                    after_idx = idx

            if before_idx is not None and after_idx is not None and before_idx > after_idx:
                result.review_signals.append(f"structural: {message}")
                finding_count += 1

    result.validation_finding_count = finding_count
    return result


def get_confidence_summary(result: RecognitionResult) -> dict:
    """Build a confidence summary from recognition result."""
    if not result.candidates:
        return {"overall": "none", "fields": {}}

    confidences = {}
    levels = {"high": 3, "medium": 2, "low": 1}

    for c in result.candidates.values():
        confidences[c.element_type] = c.confidence

    if confidences:
        avg_level = sum(levels.get(c, 0) for c in confidences.values()) / len(confidences)
        if avg_level >= 2.5:
            overall = "high"
        elif avg_level >= 1.5:
            overall = "medium"
        else:
            overall = "low"
    else:
        overall = "none"

    return {
        "overall": overall,
        "fields": confidences,
        "missing_required": result.missing_required,
        "finding_count": result.validation_finding_count,
    }
