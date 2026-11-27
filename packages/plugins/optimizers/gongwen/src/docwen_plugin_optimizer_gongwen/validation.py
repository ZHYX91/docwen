"""Post-recognition validation for gongwen documents."""

from __future__ import annotations

from typing import TYPE_CHECKING

from docwen_plugin_optimizer_gongwen.constants import MEETING_SIGNATURE_DATE

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
    positions = {
        element: index
        for index, candidate in result.candidates.items()
        if candidate.element_type != MEETING_SIGNATURE_DATE or index not in result.preserved_source_indices
        for element in candidate.represented_elements
    }

    # Check required elements
    missing = REQUIRED_ELEMENTS - positions.keys()
    result.missing_required = sorted(missing)

    # Count findings
    finding_count = len(missing)

    # Check structural rules
    for before_type, after_type, _rule_id, message in STRUCTURAL_RULES:
        if before_type in positions and after_type in positions and positions[before_type] > positions[after_type]:
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

    for index, c in result.candidates.items():
        if c.element_type == MEETING_SIGNATURE_DATE and index in result.preserved_source_indices:
            continue
        for element in c.represented_elements:
            confidences[element] = c.confidence

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
