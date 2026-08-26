from __future__ import annotations

import base64
import hashlib
import json
from dataclasses import replace
from pathlib import Path

import pytest
from jsonschema import Draft202012Validator

from docwen_core.models._resolved_numbering_semantics import (
    validate_document,
    validate_plan,
    validate_port,
)
from docwen_core.models.resolved_numbering import (
    CaptionMaterialization,
    HeadingListMaterialization,
    ResolvedNumberingPortError,
    canonicalize_numbering_plan,
    load_resolved_numbering_bytes,
)

pytestmark = pytest.mark.contract


def _digest(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


_ONE_PIXEL_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
)


def _embedded_resource(resource_id: str, content: bytes = _ONE_PIXEL_PNG) -> dict[str, object]:
    return {
        "resource_id": resource_id,
        "role": "linked_resource",
        "media_type": "image/png",
        "size_bytes": len(content),
        "sha256": hashlib.sha256(content).hexdigest(),
        "content_base64": base64.b64encode(content).decode("ascii"),
    }


def _resource_occurrence(source: str, token: str, resource_id: str, *, start: int = 0):
    source_start = source.index(token, start)
    return {
        "source_start": source_start,
        "source_end": source_start + len(token),
        "source_slice_sha256": _digest(token),
        "authored_token": token,
        "authored_locator": "same.png",
        "resource_id": resource_id,
    }


def _target(source: str, line: str, kind: str, target_id: str | None, *, level: int | None = None):
    start = source.index(line)
    return {
        "source_start": start,
        "source_end": start + len(line),
        "source_slice_sha256": _digest(line),
        "kind": kind,
        "target_id": target_id,
        "heading_level": level,
        "authored_text": line.removeprefix("# ").removeprefix("## ").removeprefix("Figure: "),
    }


def _heading_level(level: int, number_format: str, display: list[dict[str, object]], *, start: int = 1):
    return {
        "level": level,
        "start": start,
        "number_format": number_format,
        "display": display,
        "suffix": "space",
        "restart_after_level": level - 1 if level > 1 else None,
    }


def _heading_materialization(level: int) -> dict[str, object]:
    return {"type": "heading_list", "definition_id": "definition-1", "instance_id": "instance-1", "level": level}


def _plan_target(
    document_target: dict[str, object],
    *,
    enabled: bool,
    derived_number: str | None,
    materialization: dict[str, object] | None,
) -> dict[str, object]:
    return {
        "source_start": document_target["source_start"],
        "source_end": document_target["source_end"],
        "kind": document_target["kind"],
        "enabled": enabled,
        "target_id": document_target["target_id"],
        "derived_number": derived_number,
        "materialization": materialization,
    }


def _resources(
    source: str,
    targets: list[dict[str, object]],
    plan: dict[str, object],
    *,
    references: list[dict[str, object]] | None = None,
    input_id: str = "document-1",
) -> tuple[bytes, bytes]:
    source_sha = _digest(source)
    plan_sha = hashlib.sha256(canonicalize_numbering_plan(plan)).hexdigest()
    document_envelope = {
        "$schema": "urn:docwen:schema:resolved-document:v1",
        "schema": "docwen.resolved_document.v1",
        "input_id": input_id,
        "source_sha256": source_sha,
        "plan_sha256": plan_sha,
        "document": {
            "authored_markdown": source,
            "targets": targets,
            "references": references or [],
            "resource_occurrences": [],
            "citations": [],
            "resources": [],
        },
    }
    plan_envelope = {
        "$schema": "urn:docwen:schema:numbering-export-plan:v1",
        "schema": "docwen.numbering_export_plan.v1",
        "input_id": input_id,
        "source_sha256": source_sha,
        "plan_sha256": plan_sha,
        "plan": plan,
    }
    return (
        json.dumps(document_envelope, ensure_ascii=False, separators=(",", ":")).encode(),
        json.dumps(plan_envelope, ensure_ascii=False, separators=(",", ":")).encode(),
    )


def _basic_heading_resources(
    *, number_format: str = "arabic_half", start: int = 1, derived: str = "1"
) -> tuple[bytes, bytes]:
    source = "# 2.3 标题 ^1-target\n\nSee @[[#^1-target]].\n"
    heading = _target(source, "# 2.3 标题 ^1-target", "heading", "1-target", level=1)
    ref_start = source.index("@[[")
    token = "@[[#^1-target]]"
    reference = {
        "source_start": ref_start,
        "source_end": ref_start + len(token),
        "source_slice_sha256": _digest(token),
        "authored_token": token,
        "target_source_start": heading["source_start"],
        "target_source_end": heading["source_end"],
        "target_kind": "heading",
        "target_id": "1-target",
        "cached_number": derived,
        "alias": None,
    }
    plan = {
        "heading_definitions": [
            {
                "definition_id": "definition-1",
                "levels": [
                    _heading_level(
                        1,
                        number_format,
                        [{"counter": {"level": 1, "number_format": number_format}}],
                        start=start,
                    )
                ],
            }
        ],
        "heading_instances": [{"instance_id": "instance-1", "definition_id": "definition-1", "starts": []}],
        "targets": [
            _plan_target(
                heading,
                enabled=True,
                derived_number=derived,
                materialization=_heading_materialization(1),
            )
        ],
    }
    return _resources(source, [heading], plan, references=[reference])


def _cross_reference_resources(
    token: str,
    *,
    alias: str | None,
    target_kind: str = "heading",
    target_id: str | None = "1-target",
) -> tuple[bytes, bytes]:
    declaration = "# Heading ^1-target" if target_kind == "heading" else "Figure: Caption"
    source = f"{declaration}\n\n{token}\n"
    target = _target(
        source,
        declaration,
        target_kind,
        target_id,
        level=1 if target_kind == "heading" else None,
    )
    reference_start = source.index(token)
    reference = {
        "source_start": reference_start,
        "source_end": reference_start + len(token),
        "source_slice_sha256": _digest(token),
        "authored_token": token,
        "target_source_start": target["source_start"],
        "target_source_end": target["source_end"],
        "target_kind": target_kind,
        "target_id": target_id,
        "cached_number": "1",
        "alias": alias,
    }
    if target_kind == "heading":
        definitions = [
            {
                "definition_id": "definition-1",
                "levels": [
                    _heading_level(
                        1,
                        "arabic_half",
                        [{"counter": {"level": 1, "number_format": "arabic_half"}}],
                    )
                ],
            }
        ]
        instances = [{"instance_id": "instance-1", "definition_id": "definition-1", "starts": []}]
        plan_target = _plan_target(
            target,
            enabled=True,
            derived_number="1",
            materialization=_heading_materialization(1),
        )
    else:
        definitions = []
        instances = []
        plan_target = _plan_target(target, enabled=False, derived_number=None, materialization=None)
    plan = {
        "heading_definitions": definitions,
        "heading_instances": instances,
        "targets": [plan_target],
    }
    return _resources(source, [target], plan, references=[reference])


def _caption_matrix_resources(
    materialization_type: str,
    action: str,
    *,
    chapter_cache: str = "1",
    chapter_level: int = 1,
    sequence_start: int = 1,
) -> tuple[bytes, bytes]:
    source = "# Chapter\n## Scope\nFigure: Alpha\n"
    chapter = _target(source, "# Chapter", "heading", "chapter", level=1)
    scope = _target(source, "## Scope", "heading", "scope", level=2)
    figure = _target(source, "Figure: Alpha", "figure", "figure-1")
    targets = [chapter, scope, figure]
    if action == "continue":
        start_value = None
        sequence_value = 1
        restart_level = None
        restart_style = None
    elif action == "reset_to_start":
        start_value = sequence_start
        sequence_value = sequence_start
        restart_level = None
        restart_style = None
    else:
        start_value = 1
        sequence_value = 1
        restart_level = 2
        restart_style = "heading_2"
    sequence_cache = str(sequence_value)
    if materialization_type == "chapter_seq":
        chapter_heading_level = chapter_level
        chapter_heading_style = f"heading_{chapter_level}"
        separator = "-"
        total = f"{chapter_cache}-{sequence_cache}"
        cached_chapter: str | None = chapter_cache
    else:
        chapter_heading_level = None
        chapter_heading_style = None
        separator = None
        total = sequence_cache
        cached_chapter = None
    caption_materialization = {
        "type": materialization_type,
        "counter": "Figure",
        "number_format": "arabic_half",
        "sequence_action": action,
        "start_value": start_value,
        "restart_heading_level": restart_level,
        "restart_heading_style": restart_style,
        "chapter_heading_level": chapter_heading_level,
        "chapter_heading_style": chapter_heading_style,
        "chapter_separator": separator,
        "chapter_cached_number": cached_chapter,
        "sequence_cached_number": sequence_cache,
        "localized_label": "Figure",
        "label_separator": " ",
    }
    plan = {
        "heading_definitions": [
            {
                "definition_id": "definition-1",
                "levels": [
                    _heading_level(
                        1,
                        "arabic_half",
                        [{"counter": {"level": 1, "number_format": "arabic_half"}}],
                    ),
                    _heading_level(
                        2,
                        "arabic_half",
                        [
                            {"counter": {"level": 1, "number_format": "arabic_half"}},
                            {"literal": "-"},
                            {"counter": {"level": 2, "number_format": "arabic_half"}},
                        ],
                    ),
                ],
            }
        ],
        "heading_instances": [{"instance_id": "instance-1", "definition_id": "definition-1", "starts": []}],
        "targets": [
            _plan_target(
                chapter,
                enabled=True,
                derived_number="1",
                materialization=_heading_materialization(1),
            ),
            _plan_target(
                scope,
                enabled=True,
                derived_number="1-1",
                materialization=_heading_materialization(2),
            ),
            _plan_target(
                figure,
                enabled=True,
                derived_number=total,
                materialization=caption_materialization,
            ),
        ],
    }
    return _resources(source, targets, plan)


def _disabled_heading_scope_resources(*, chapter_seq: bool) -> tuple[bytes, bytes]:
    source = "# Numbered\nFigure: One\n# Disabled\nFigure: Two\n"
    first_heading = _target(source, "# Numbered", "heading", "heading-1", level=1)
    first_figure = _target(source, "Figure: One", "figure", "figure-1")
    second_heading = _target(source, "# Disabled", "heading", "heading-2", level=1)
    second_figure = _target(source, "Figure: Two", "figure", "figure-2")

    def caption_materialization() -> dict[str, object]:
        return {
            "type": "chapter_seq" if chapter_seq else "simple_seq",
            "counter": "Figure",
            "number_format": "arabic_half",
            "sequence_action": "restart_by_heading_level",
            "start_value": 1,
            "restart_heading_level": 1,
            "restart_heading_style": "heading_1",
            "chapter_heading_level": 1 if chapter_seq else None,
            "chapter_heading_style": "heading_1" if chapter_seq else None,
            "chapter_separator": "-" if chapter_seq else None,
            "chapter_cached_number": "1" if chapter_seq else None,
            "sequence_cached_number": "1",
            "localized_label": "Figure",
            "label_separator": " ",
        }

    derived = "1-1" if chapter_seq else "1"
    targets = [first_heading, first_figure, second_heading, second_figure]
    plan = {
        "heading_definitions": [
            {
                "definition_id": "definition-1",
                "levels": [
                    _heading_level(
                        1,
                        "arabic_half",
                        [{"counter": {"level": 1, "number_format": "arabic_half"}}],
                    )
                ],
            }
        ],
        "heading_instances": [{"instance_id": "instance-1", "definition_id": "definition-1", "starts": []}],
        "targets": [
            _plan_target(
                first_heading,
                enabled=True,
                derived_number="1",
                materialization=_heading_materialization(1),
            ),
            _plan_target(
                first_figure,
                enabled=True,
                derived_number=derived,
                materialization=caption_materialization(),
            ),
            _plan_target(
                second_heading,
                enabled=False,
                derived_number=None,
                materialization=None,
            ),
            _plan_target(
                second_figure,
                enabled=True,
                derived_number=derived,
                materialization=caption_materialization(),
            ),
        ],
    }
    return _resources(source, targets, plan)


__all__ = (
    "_ONE_PIXEL_PNG",
    "CaptionMaterialization",
    "Draft202012Validator",
    "HeadingListMaterialization",
    "Path",
    "ResolvedNumberingPortError",
    "_basic_heading_resources",
    "_caption_matrix_resources",
    "_cross_reference_resources",
    "_digest",
    "_disabled_heading_scope_resources",
    "_embedded_resource",
    "_plan_target",
    "_resource_occurrence",
    "_resources",
    "_target",
    "base64",
    "canonicalize_numbering_plan",
    "hashlib",
    "json",
    "load_resolved_numbering_bytes",
    "pytest",
    "pytestmark",
    "replace",
    "validate_document",
    "validate_plan",
    "validate_port",
)
