"""Validation-driven re-evaluation of unique element assignments.

When the validator finds missing required fields, try switching or dropping
low-confidence unique element assignments to see if finding count decreases.
"""

from __future__ import annotations

from docwen_plugin_optimizer_gongwen.models import (
    ParagraphFeature,
    RecognitionCandidate,
    RecognitionResult,
)


def maybe_reevaluate(
    scorer,
    features: list[ParagraphFeature],
    result: RecognitionResult,
) -> RecognitionResult:
    """Re-evaluate unique element assignments if validation finds issues.

    If required fields (title, issuing_authority_signature, issue_date) are
    missing, try both drop and switch modes to improve the result.

    Modes:
    - drop: Remove a low-confidence assignment and re-score for missing types.
    - switch: Replace a low-confidence assignment with its runner-up candidate
      and check if validation improves.

    Returns the best result found, or the original if no improvement.
    """
    required_types = {"title", "issuing_authority_signature", "issue_date"}
    found_types = {c.element_type for c in result.candidates.values()}
    missing = required_types - found_types

    if not missing:
        result.missing_required = []
        return result

    result.missing_required = sorted(missing)

    # Find low-confidence candidates that might be misclassified
    low_conf_indices = [idx for idx, c in result.candidates.items() if c.confidence == "low"]

    if not low_conf_indices:
        return result

    best_result = result
    best_missing_count = len(missing)

    # ── 1. DROP mode: drop low-confidence assignment, re-score for missing types
    for idx in low_conf_indices:
        dropped_type = result.candidates[idx].element_type
        modified = dict(result.candidates)
        del modified[idx]

        pf = features[idx]
        all_candidates = scorer.score_round(pf, round_group="round1")
        for cand in all_candidates:
            if cand.element_type in missing:
                modified[idx] = cand
                new_missing = missing - {cand.element_type}
                if len(new_missing) < best_missing_count:
                    best_result = _build_result(
                        modified,
                        result,
                        new_missing,
                        f"Re-evaluated(drop): para_{idx} {dropped_type}->{cand.element_type}",
                    )
                    best_missing_count = len(new_missing)
                break  # use first match

    # ── 2. SWITCH mode: replace with runner-up element
    unique_element_types = {
        "title",
        "issuing_authority_mark",
        "doc_number",
        "security",
        "urgency",
        "copy_id",
        "combined_id",
        "combined_doc_number_signer",
        "issuing_authority_signature",
        "issue_date",
        "printing_date",
        "recipient",
        "notes",
        "disclosure",
        "copy_to",
        "printing_authority",
        "attachment_header",
    }

    for idx in low_conf_indices:
        runner_up = _get_runner_up(idx, scorer, features, result)
        if runner_up is None:
            continue

        # Skip if runner-up type would create a duplicate for unique elements
        if runner_up.element_type in unique_element_types and runner_up.element_type in found_types:
            continue  # unique element types shouldn't appear twice

        modified = dict(result.candidates)
        modified[idx] = runner_up
        new_found = {c.element_type for c in modified.values()}
        new_missing = required_types - new_found

        if len(new_missing) < best_missing_count:
            best_result = _build_result(
                modified, result, new_missing, f"Re-evaluated(switch): para_{idx} ->{runner_up.element_type}"
            )
            best_missing_count = len(new_missing)

    return best_result


def _get_runner_up(
    para_idx: int,
    scorer,
    features: list[ParagraphFeature],
    result: RecognitionResult,
) -> RecognitionCandidate | None:
    """Get the highest-scoring alternative candidate for a paragraph."""
    pf = features[para_idx]
    all_candidates = scorer.score_round(pf, round_group="round1")

    # Filter out already-occupied unique element types
    occupied = {c.element_type for i, c in result.candidates.items() if i != para_idx}

    unique_element_types = {
        "title",
        "issuing_authority_mark",
        "doc_number",
        "security",
        "urgency",
        "copy_id",
        "combined_id",
        "combined_doc_number_signer",
        "issuing_authority_signature",
        "issue_date",
        "printing_date",
        "recipient",
        "notes",
        "disclosure",
        "copy_to",
        "printing_authority",
        "attachment_header",
    }
    available = [
        c for c in all_candidates if c.element_type not in occupied or c.element_type not in unique_element_types
    ]

    if len(available) >= 2:
        runner_up = available[1]
        if runner_up.score >= 80:  # minimum threshold for switch
            return runner_up
    return None


def _build_result(
    modified_candidates: dict,
    original: RecognitionResult,
    missing: set,
    signal: str,
) -> RecognitionResult:
    """Build a RecognitionResult from modified candidates."""
    return RecognitionResult(
        candidates=modified_candidates,
        yaml_info=original.yaml_info,
        skip_indices=original.skip_indices,
        review_signals=[*original.review_signals, signal],
        missing_required=sorted(missing),
        validation_finding_count=len(missing),
    )
