"""Admission gates for CommonMark Heading levels versus Word definitions."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pytest
from jsonschema import Draft202012Validator

from docwen_core.models.resolved_numbering import (
    ResolvedNumberingPortError,
    canonicalize_numbering_plan,
    load_resolved_numbering_bytes,
)

pytestmark = pytest.mark.contract


def _sha(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def _document_target(source: str, *, kind: str, heading_level: int | None) -> dict[str, object]:
    line = source.rstrip("\n")
    return {
        "source_start": 0,
        "source_end": len(line),
        "source_slice_sha256": _sha(line),
        "kind": kind,
        "target_id": None,
        "heading_level": heading_level,
        "authored_text": line.lstrip("# ").removeprefix("Figure: "),
    }


def _plan_target(
    document_target: dict[str, object],
    materialization: dict[str, object],
    *,
    derived_number: str,
) -> dict[str, object]:
    return {
        "source_start": document_target["source_start"],
        "source_end": document_target["source_end"],
        "kind": document_target["kind"],
        "enabled": True,
        "target_id": document_target["target_id"],
        "derived_number": derived_number,
        "materialization": materialization,
    }


def _resources(
    source: str,
    document_target: dict[str, object],
    plan: dict[str, object],
) -> tuple[bytes, bytes]:
    plan_sha256 = hashlib.sha256(canonicalize_numbering_plan(plan)).hexdigest()
    document = {
        "$schema": "urn:docwen:schema:resolved-document:v1",
        "schema": "docwen.resolved_document.v1",
        "input_id": "heading-level-gate",
        "source_sha256": _sha(source),
        "plan_sha256": plan_sha256,
        "document": {
            "authored_markdown": source,
            "targets": [document_target],
            "references": [],
            "resource_occurrences": [],
            "citations": [],
            "resources": [],
        },
    }
    plan_envelope = {
        "$schema": "urn:docwen:schema:numbering-export-plan:v1",
        "schema": "docwen.numbering_export_plan.v1",
        "input_id": "heading-level-gate",
        "source_sha256": _sha(source),
        "plan_sha256": plan_sha256,
        "plan": plan,
    }
    return (
        json.dumps(document, separators=(",", ":")).encode(),
        json.dumps(plan_envelope, separators=(",", ":")).encode(),
    )


def _heading_level_definition(level: int) -> dict[str, object]:
    return {
        "level": level,
        "start": 1,
        "number_format": "arabic_half",
        "display": [{"counter": {"level": level, "number_format": "arabic_half"}}],
        "suffix": "space",
        "restart_after_level": level - 1 if level > 1 else None,
    }


def _heading_plan(document_target: dict[str, object], *, used_level: int) -> dict[str, object]:
    return {
        "heading_definitions": [
            {
                "definition_id": "definition-1",
                "levels": [_heading_level_definition(level) for level in range(1, 10)],
            }
        ],
        "heading_instances": [
            {
                "instance_id": "instance-1",
                "definition_id": "definition-1",
                "starts": [{"level": 9, "value": 2}],
            }
        ],
        "targets": [
            _plan_target(
                document_target,
                {
                    "type": "heading_list",
                    "definition_id": "definition-1",
                    "instance_id": "instance-1",
                    "level": used_level,
                },
                derived_number="2" if used_level == 9 else "1",
            )
        ],
    }


def _caption_plan(
    document_target: dict[str, object],
    *,
    binding: str,
    level: int,
) -> dict[str, object]:
    chapter = binding == "chapter"
    materialization = {
        "type": "chapter_seq" if chapter else "simple_seq",
        "counter": "Figure",
        "number_format": "arabic_half",
        "sequence_action": "continue" if chapter else "restart_by_heading_level",
        "start_value": None if chapter else 1,
        "restart_heading_level": None if chapter else level,
        "restart_heading_style": None if chapter else f"heading_{level}",
        "chapter_heading_level": level if chapter else None,
        "chapter_heading_style": f"heading_{level}" if chapter else None,
        "chapter_separator": "-" if chapter else None,
        "chapter_cached_number": "1" if chapter else None,
        "sequence_cached_number": "1",
        "localized_label": "Figure",
        "label_separator": " ",
    }
    return {
        "heading_definitions": [],
        "heading_instances": [],
        "targets": [
            _plan_target(
                document_target,
                materialization,
                derived_number="1-1" if chapter else "1",
            )
        ],
    }


@pytest.mark.parametrize("level", [7, 8, 9])
def test_authored_markdown_heading_targets_through_nine_are_valid(level: int) -> None:
    source = f"{'#' * level} Heading\n"
    document_target = _document_target(source, kind="heading", heading_level=level)
    plan = _heading_plan(document_target, used_level=level)

    port = load_resolved_numbering_bytes(*_resources(source, document_target, plan))

    assert port.document.targets[0].heading_level == level
    assert getattr(port.plan.targets[0].materialization, "level", None) == level


def test_authored_heading_and_used_heading_list_level_six_are_accepted() -> None:
    source = "###### Heading\n"
    document_target = _document_target(source, kind="heading", heading_level=6)
    plan = _heading_plan(document_target, used_level=6)

    port = load_resolved_numbering_bytes(*_resources(source, document_target, plan))

    assert port.document.targets[0].heading_level == 6
    materialization = port.plan.targets[0].materialization
    assert materialization is not None
    assert getattr(materialization, "level", None) == 6


def test_used_heading_list_level_ten_is_invalid() -> None:
    source = "# Heading\n"
    document_target = _document_target(source, kind="heading", heading_level=1)
    plan = _heading_plan(document_target, used_level=10)

    with pytest.raises(ResolvedNumberingPortError) as rejected:
        load_resolved_numbering_bytes(*_resources(source, document_target, plan))

    assert rejected.value.code == "docwen.numbering_export_plan.invalid"


@pytest.mark.parametrize("binding", ["chapter", "restart"])
def test_used_caption_heading_binding_level_ten_is_invalid(binding: str) -> None:
    source = "Figure: Caption\n"
    document_target = _document_target(source, kind="figure", heading_level=None)
    plan = _caption_plan(document_target, binding=binding, level=10)

    with pytest.raises(ResolvedNumberingPortError) as rejected:
        load_resolved_numbering_bytes(*_resources(source, document_target, plan))

    assert rejected.value.code == "docwen.numbering_export_plan.invalid"


def test_word_definition_and_instance_level_nine_remain_valid_when_unused() -> None:
    source = "# Heading\n"
    document_target = _document_target(source, kind="heading", heading_level=1)
    plan = _heading_plan(document_target, used_level=1)

    port = load_resolved_numbering_bytes(*_resources(source, document_target, plan))

    assert port.plan.heading_definitions[0].levels[-1].level == 9
    assert port.plan.heading_instances[0].starts[-1].level == 9


def test_conformance_level_fixtures_match_the_narrowed_json_schemas() -> None:
    root = Path(__file__).parents[3]
    contracts = root / "contracts"
    document_validator = Draft202012Validator(
        json.loads((contracts / "schemas/docwen.resolved_document.v1.schema.json").read_text(encoding="utf-8"))
    )
    plan_validator = Draft202012Validator(
        json.loads((contracts / "schemas/docwen.numbering_export_plan.v1.schema.json").read_text(encoding="utf-8"))
    )
    invalid_documents = ["resolved-document.heading-level-10.json"]
    invalid_plans = [
        "numbering-export-plan.heading-list-level-10.json",
        "numbering-export-plan.restart-level-10.json",
        "numbering-export-plan.chapter-level-10.json",
    ]

    for name in invalid_documents:
        payload = json.loads((contracts / "fixtures/invalid" / name).read_text(encoding="utf-8"))
        assert not document_validator.is_valid(payload)
    for name in invalid_plans:
        payload = json.loads((contracts / "fixtures/invalid" / name).read_text(encoding="utf-8"))
        assert not plan_validator.is_valid(payload)
        materialization = payload["plan"]["targets"][0]["materialization"]
        if materialization["type"] == "heading_list":
            materialization["level"] = 9
        elif materialization["type"] == "chapter_seq":
            materialization["chapter_heading_level"] = 9
            materialization["chapter_heading_style"] = "heading_9"
        else:
            materialization["restart_heading_level"] = 9
            materialization["restart_heading_style"] = "heading_9"
        plan_validator.validate(payload)

    valid_document, valid_plan = _resources(
        "# Heading\n",
        _document_target("# Heading\n", kind="heading", heading_level=1),
        _heading_plan(_document_target("# Heading\n", kind="heading", heading_level=1), used_level=1),
    )
    document_validator.validate(json.loads(valid_document))
    plan_validator.validate(json.loads(valid_plan))


def test_docs_expose_the_docwen_heading_extension_through_level_nine() -> None:
    root = Path(__file__).parents[3]
    structured = (root / "docs/specs/structured-numbering-phases.md").read_text(encoding="utf-8")
    machine = (root / "docs/specs/machine-protocol-v1.md").read_text(encoding="utf-8")
    markdown = (root / "docs/specs/markdown-compatibility.md").read_text(encoding="utf-8")
    structured = " ".join(structured.split())
    machine = " ".join(machine.split())
    markdown = " ".join(markdown.split())

    assert "DocWen ATX extension levels 7..9" in structured
    assert "every used `heading_list` binding" in structured
    assert "caption restart/chapter binding are restricted to 1..9" in structured
    assert "DocWen ATX extension levels 7..9" in machine
    assert "levels 7..9 are DocWen extensions" in markdown
