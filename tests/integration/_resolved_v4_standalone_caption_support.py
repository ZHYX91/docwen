"""Synthetic resolved-v4 standalone-caption resources for readback tests."""

from __future__ import annotations

import hashlib
import json
from dataclasses import asdict
from pathlib import Path

from docwen_core.models.resolved_numbering import (
    CaptionMaterialization,
    NumberingTarget,
    ResolvedDocument,
    ResolvedDocumentTarget,
    ResolvedNumberingPlan,
    ResolvedReference,
    canonicalize_numbering_plan,
)


def _sha(value: str) -> str:
    return hashlib.sha256(value.encode()).hexdigest()


def _write_standalone_caption_pair(tmp_path: Path) -> tuple[Path, Path]:
    source = (
        "Figure: Planned figure ^fig-plan\n\n\nBody A.\n\n"
        "Table: Planned table\n\n\nBody B.\n\n"
        "Code: Planned code\n\n\nBody C.\n\n"
        "See @[[#^fig-plan]].\n"
    )
    figure_start = source.index("Figure:")
    figure_end = source.index("\n", figure_start)
    table_start = source.index("Table:")
    table_end = source.index("\n", table_start)
    code_start = source.index("Code:")
    code_end = source.index("\n", code_start)
    targets = (
        ResolvedDocumentTarget(
            figure_start,
            figure_end,
            _sha(source[figure_start:figure_end]),
            "figure",
            "fig-plan",
            None,
            "Planned figure",
        ),
        ResolvedDocumentTarget(
            table_start,
            table_end,
            _sha(source[table_start:table_end]),
            "table",
            None,
            None,
            "Planned table",
        ),
        ResolvedDocumentTarget(
            code_start,
            code_end,
            _sha(source[code_start:code_end]),
            "code_block",
            None,
            None,
            "Planned code",
        ),
    )
    reference_token = "@[[#^fig-plan]]"
    reference_start = source.index(reference_token)
    reference = ResolvedReference(
        reference_start,
        reference_start + len(reference_token),
        _sha(reference_token),
        reference_token,
        figure_start,
        figure_end,
        "figure",
        "fig-plan",
        "1",
        None,
    )
    figure_materialization = CaptionMaterialization(
        "simple_seq",
        "Figure",
        "arabic_half",
        "continue",
        None,
        None,
        None,
        None,
        None,
        None,
        None,
        "1",
        "Figure",
        " ",
    )
    table_materialization = CaptionMaterialization(
        "simple_seq",
        "Table",
        "arabic_half",
        "continue",
        None,
        None,
        None,
        None,
        None,
        None,
        None,
        "1",
        "Table",
        " ",
    )
    plan = ResolvedNumberingPlan(
        (),
        (),
        (
            NumberingTarget(figure_start, figure_end, "figure", True, "fig-plan", "1", figure_materialization),
            NumberingTarget(table_start, table_end, "table", True, None, "1", table_materialization),
            NumberingTarget(code_start, code_end, "code_block", False, None, None, None),
        ),
    )
    plan_body = json.loads(json.dumps(asdict(plan)))
    plan_sha256 = hashlib.sha256(canonicalize_numbering_plan(plan_body)).hexdigest()
    source_sha256 = _sha(source)
    document = ResolvedDocument(source, targets, (reference,), (), (), ())
    neutral_payload = {
        "$schema": "urn:docwen:schema:resolved-document:v1",
        "schema": "docwen.resolved_document.v1",
        "input_id": "standalone-captions",
        "source_sha256": source_sha256,
        "plan_sha256": plan_sha256,
        "document": asdict(document),
    }
    plan_payload = {
        "$schema": "urn:docwen:schema:numbering-export-plan:v1",
        "schema": "docwen.numbering_export_plan.v1",
        "input_id": "standalone-captions",
        "source_sha256": source_sha256,
        "plan_sha256": plan_sha256,
        "plan": plan_body,
    }
    neutral = tmp_path / "standalone-neutral.json"
    plan_path = tmp_path / "standalone-plan.json"
    neutral.write_text(json.dumps(neutral_payload, separators=(",", ":")), encoding="utf-8")
    plan_path.write_text(json.dumps(plan_payload, separators=(",", ":")), encoding="utf-8")
    return neutral, plan_path
