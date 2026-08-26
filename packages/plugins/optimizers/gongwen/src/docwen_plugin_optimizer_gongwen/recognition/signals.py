"""Structured recognition review signals for close matches and low confidence."""

from __future__ import annotations

from typing import TYPE_CHECKING

from docwen_plugin_optimizer_gongwen.models import RecognitionResult

if TYPE_CHECKING:
    from docwen_plugin_optimizer_gongwen.models import ParagraphFeature

# ── Constants ──
CLOSE_UNIQUE_MATCH_GAP_THRESHOLD = 20


def collect_structured_signals(
    result: RecognitionResult,
    scorer,  # ElementScorer for context queries
    features: list[ParagraphFeature],
) -> dict:
    """Collect structured review signals from recognition result.

    Returns dict with keys:
        close_unique_matches:    list[{paragraph_index, element, round, score,
                                      runner_up_element, runner_up_score, score_gap}]
        used_fallback:           list[{paragraph_index, element, round, reason}]
        low_confidence_passes:   list[{paragraph_index, element, round, score, threshold}]
        needs_review_reasons:    list[str] — union of above signal type names
        gongwen_warnings:        list[{code, severity, scope, message, details}]
        missing_required:        list[str]
        recognition_summary:     dict
    """
    signals: dict = {
        "close_unique_matches": [],
        "used_fallback": [],
        "low_confidence_passes": [],
        "needs_review_reasons": [],
    }

    # --- Detect close unique matches (score gap <= 20) ---
    for idx, candidate in result.candidates.items():
        # Re-score to get all candidates trace
        pf = features[idx] if idx < len(features) else None
        if pf is None:
            continue
        all_cands = scorer.score_round(pf, round_group="round1")
        if len(all_cands) >= 2:
            gap = all_cands[0].score - all_cands[1].score
            if gap <= CLOSE_UNIQUE_MATCH_GAP_THRESHOLD:
                signals["close_unique_matches"].append(
                    {
                        "paragraph_index": idx,
                        "element": candidate.element_type,
                        "round": 1,
                        "score": all_cands[0].score,
                        "runner_up_element": all_cands[1].element_type,
                        "runner_up_score": all_cands[1].score,
                        "score_gap": gap,
                    }
                )

    # --- Detect low confidence passes ---
    for idx, candidate in result.candidates.items():
        if candidate.confidence == "low":
            signals["low_confidence_passes"].append(
                {
                    "paragraph_index": idx,
                    "element": candidate.element_type,
                    "round": 1,
                    "score": candidate.score,
                    "threshold": 80,
                }
            )

    # --- Build needs_review_reasons ---
    if signals["close_unique_matches"]:
        signals["needs_review_reasons"].append("close_unique_match")
    if signals["used_fallback"]:
        signals["needs_review_reasons"].append("used_fallback")
    if signals["low_confidence_passes"]:
        signals["needs_review_reasons"].append("low_confidence_pass")

    # --- Build gongwen_warnings ---
    gongwen_warnings: list[dict] = []
    if result.missing_required:
        gongwen_warnings.append(
            {
                "code": "GW001",
                "severity": "error",
                "scope": "recognition",
                "message": f"缺少必需字段: {', '.join(result.missing_required)}",
                "details": {"missing": result.missing_required},
            }
        )
    risky_tables = [feature for feature in features if feature.is_table_anchor and feature.table_fidelity_risks]
    if risky_tables:
        risky_table_indices = {feature.table_index for feature in risky_tables if feature.table_index is not None}
        table_feature_count = sum(feature.table_index in risky_table_indices for feature in features)
        fidelity_risks = sorted({risk for feature in risky_tables for risk in feature.table_fidelity_risks})
        rendered_table_count = sum(
            feature.table_output_mode == "rendered"
            or (not feature.table_output_mode and feature.index not in result.skip_indices)
            for feature in risky_tables
        )
        structural_table_count = sum(feature.table_output_mode == "structural_metadata" for feature in risky_tables)
        unrepresentable_table_count = sum(feature.table_output_mode == "unrepresentable" for feature in risky_tables)
        if structural_table_count and not rendered_table_count and not unrepresentable_table_count:
            table_message = "结构性表格已提取为公文字段并从正文移除，但其中部分富内容无法写入 YAML，需人工核对。"
        elif unrepresentable_table_count and not rendered_table_count and not structural_table_count:
            table_message = "表格不含可生成的纯文本，但含有无法原样表达的内容，需人工核对。"
        else:
            table_message = (
                "表格中可表示的行列和合并关系已保留；结构性字段已提取，但部分单元格内容无法原样表达，需人工核对。"
            )
        gongwen_warnings.append(
            {
                "code": "GW002",
                "severity": "warning",
                "scope": "table_semantics",
                "message": table_message,
                "details": {
                    "table_count": len(risky_table_indices),
                    "table_cell_feature_count": table_feature_count,
                    "fidelity_risks": fidelity_risks,
                    "rendered_table_count": rendered_table_count,
                    "structural_table_count": structural_table_count,
                    "unrepresentable_table_count": unrepresentable_table_count,
                },
            }
        )

    # --- Build recognition_summary ---
    needs_review = bool(gongwen_warnings or signals["needs_review_reasons"] or result.missing_required)
    recognition_summary = {
        "status": "needs_review" if needs_review else "ok",
        "needs_review": needs_review,
        "recognized_paragraph_count": len(result.candidates),
        "warning_count": len(gongwen_warnings),
        "missing_required_element_count": len(result.missing_required),
    }

    return {
        **signals,
        "gongwen_warnings": gongwen_warnings,
        "missing_required": result.missing_required,
        "recognition_summary": recognition_summary,
    }
