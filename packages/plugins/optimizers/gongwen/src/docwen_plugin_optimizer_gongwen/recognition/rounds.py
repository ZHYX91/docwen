"""Three-round element recognition orchestration.

The rounds enforce these matching rules:
- Round 1: unique elements, processed paragraph-by-paragraph in document order.
  Each paragraph competes for all remaining unique round1 types; the best
  above-threshold match wins and the type is removed from the pool.
- Round 2: unique elements, same flow but for round2 types.
- Round 3: non-unique elements, scored for every unassigned paragraph.
"""

from __future__ import annotations

from docwen_plugin_optimizer_gongwen.models import (
    ParagraphFeature,
    RecognitionCandidate,
    RecognitionResult,
)


def run_unique_rounds(
    scorer,
    features: list[ParagraphFeature],
) -> RecognitionResult:
    """Run rounds 1 and 2 (unique element types).

    Faithfully mirrors the old paragraph-by-paragraph scoring loop:
    each paragraph is scored against all remaining unique types; the
    highest-scoring match above threshold wins the type for that paragraph.

    Returns a RecognitionResult with all assigned candidates.
    """
    scorer.reset_context(features)
    candidates: dict[int, RecognitionCandidate] = {}
    identified_unique: set[str] = set()

    # ── Round 1 ─────────────────────────────────────────────────
    for pf in features:
        if pf.index in candidates:
            continue

        best = _score_best_candidate(
            scorer,
            pf,
            round_group="round1",
            skip_types=identified_unique,
        )
        if best is not None:
            candidates[pf.index] = best
            identified_unique.add(best.element_type)
            scorer.update_context(best.element_type, pf.index)

    # ── Round 2 ─────────────────────────────────────────────────
    for pf in features:
        if pf.index in candidates:
            continue

        best = _score_best_candidate(
            scorer,
            pf,
            round_group="round2",
            skip_types=identified_unique,
        )
        if best is not None:
            candidates[pf.index] = best
            identified_unique.add(best.element_type)
            scorer.update_context(best.element_type, pf.index)

    return RecognitionResult(
        candidates=candidates,
        yaml_info={},
        skip_indices=[],
        review_signals=[],
        missing_required=[],
        validation_finding_count=0,
    )


def run_rounds(
    scorer,
    features: list[ParagraphFeature],
) -> RecognitionResult:
    """Run all three rounds of element recognition.

    Rounds 1-2: unique elements (one paragraph per element type).
    Round 3: non-unique elements (multiple paragraphs can be body, etc.).
    """
    result = run_unique_rounds(scorer, features)
    result = run_round3(scorer, features, result)
    return result


def run_round3(
    scorer,
    features: list[ParagraphFeature],
    result: RecognitionResult,
) -> RecognitionResult:
    """Run round 3 for non-unique element types.

    Body, attachment items, signer following, etc. can appear multiple times.
    Unassigned paragraphs are scored against all round3 types; the best
    above-threshold match is assigned.
    """
    for pf in features:
        if pf.index in result.candidates:
            continue

        best = _score_best_candidate(
            scorer,
            pf,
            round_group="round3",
            skip_types=set(),  # non-unique types are always available
        )
        if best is not None:
            result.candidates[pf.index] = best
            scorer.update_context(best.element_type, pf.index)
        else:
            # Default to body for unassigned paragraphs with Chinese text
            from docwen_plugin_optimizer_gongwen.utils import contains_chinese

            if contains_chinese(pf.text) and len(pf.text.strip()) > 2:
                result.candidates[pf.index] = RecognitionCandidate(
                    element_type="body",
                    score=60,  # NON_UNIQUE_SCORE_THRESHOLD
                    para_index=pf.index,
                    trace=["default_body"],
                    confidence="low",
                )
                scorer.update_context("body", pf.index)

    return result


def _score_best_candidate(
    scorer,
    pf: ParagraphFeature,
    *,
    round_group: str = "round1",
    skip_types: set[str] | None = None,
) -> RecognitionCandidate | None:
    """Score a paragraph against all types in a round group and return the best.

    Args:
        scorer: ElementScorer instance.
        pf: Paragraph to score.
        round_group: "round1", "round2", or "round3".
        skip_types: Set of element types to skip (already assigned unique types).

    Returns:
        The highest-scoring RecognitionCandidate, or None if no match above threshold.
    """
    candidates = scorer.score_round(pf, round_group=round_group)

    if skip_types:
        candidates = [c for c in candidates if c.element_type not in skip_types]

    if not candidates:
        return None

    # candidates are already sorted by score descending in score_round
    return candidates[0]
