"""Closed input and source-projection seam for resolved v4 MD→DOCX.

This module is deliberately smaller than the converter.  It owns the two
security-sensitive transitions that must happen before the ordinary Markdown
pipeline is allowed to run:

* re-admit the exact ``neutral_document`` + ``numbering_export_plan`` pair;
* compose typed semantic markers and embedded-resource replacements by their
  authenticated source coordinates.

It never resolves an authored locator, searches for a token, or falls back to
the neutral JSON file's directory.
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from docwen_core.models.resolved_numbering import (
    NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
    RESOLVED_DOCUMENT_MEDIA_TYPE,
    ResolvedNumberingPort,
    ResolvedNumberingPortError,
    load_resolved_numbering_port,
)
from docwen_core.resolved_resource_staging import ResolvedResourceBinding
from docwen_plugin_markdown.resolved_runtime_v4 import (
    ResolvedRuntimeV4Plan,
    ResolvedRuntimeV4Unsupported,
    prepare_resolved_runtime_v4,
)
from docwen_plugin_markdown.resolved_source_carriers_v4 import (
    ResolvedSourceCarrierPlanV4,
    prepare_resolved_source_carriers_v4,
)
from docwen_plugin_markdown.runtime_semantics_v3 import RuntimeSemanticsV3Unsupported

_NEUTRAL_ROLE = "neutral_document"
_PLAN_ROLE = "numbering_export_plan"
_RESOLVED_ROLES = frozenset({_NEUTRAL_ROLE, _PLAN_ROLE})
_CAPTION_CHILDREN_KEY = "_docwen_resolved_v4_caption_children"


class ResolvedConversionV4Unsupported(ValueError):
    """The request cannot enter the closed resolved-v4 conversion route."""

    def __init__(self, code: str, message: str) -> None:
        self.code = code
        super().__init__(message)


@dataclass(frozen=True, slots=True)
class PreparedResolvedInputsV4:
    """One immutable typed port and its source-marker projection."""

    port: ResolvedNumberingPort
    runtime_plan: ResolvedRuntimeV4Plan
    source_carrier_plan: ResolvedSourceCarrierPlanV4
    neutral_document_path: Path
    numbering_export_plan_path: Path


@dataclass(frozen=True, slots=True)
class ProjectedResolvedMarkdownV4:
    """Markdown ready for ordinary parsing plus its closed image inventory."""

    markdown: str
    expected_image_urls: tuple[str, ...]
    runtime_plan: ResolvedRuntimeV4Plan
    source_carrier_plan: ResolvedSourceCarrierPlanV4


@dataclass(frozen=True, slots=True)
class _SourceEdit:
    source_start: int
    source_end: int
    original: str
    replacement: str
    role: str


def claims_resolved_v4_inputs(resources: tuple[Any, ...]) -> bool:
    """Return whether any input claims the v4 route.

    A partial pair still returns ``True`` so callers fail closed here instead
    of falling through to the historical source-file converter.
    """

    return any(getattr(item, "input_role", None) in _RESOLVED_ROLES for item in resources)


def load_resolved_v4_inputs(workspace: Any) -> PreparedResolvedInputsV4:
    """Re-admit the exact typed pair from request-owned Workspace copies."""

    resources = tuple(workspace.input_resources())
    roles = [getattr(item, "input_role", None) for item in resources]
    if not claims_resolved_v4_inputs(resources):
        raise ResolvedConversionV4Unsupported(
            "docwen.numbering_export_plan.missing",
            "the request does not declare a resolved-document numbering input pair",
        )
    if len(resources) != 2 or set(roles) != _RESOLVED_ROLES or len(set(roles)) != 2:
        code = "docwen.numbering_export_plan.missing" if _PLAN_ROLE not in roles else "docwen.resolved_document.invalid"
        raise ResolvedConversionV4Unsupported(code, "resolved v4 conversion requires one exact input pair")

    neutral = resources[roles.index(_NEUTRAL_ROLE)]
    plan = resources[roles.index(_PLAN_ROLE)]
    _prove_input_ref(
        neutral,
        role=_NEUTRAL_ROLE,
        kind="document",
        media_type=RESOLVED_DOCUMENT_MEDIA_TYPE,
    )
    _prove_input_ref(
        plan,
        role=_PLAN_ROLE,
        kind="resource",
        media_type=NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
    )
    neutral_path = Path(neutral.path)
    plan_path = Path(plan.path)
    try:
        port = load_resolved_numbering_port(neutral_path, plan_path)
        runtime_plan = prepare_resolved_runtime_v4(port)
        source_carrier_plan = prepare_resolved_source_carriers_v4(
            port.document.authored_markdown,
            input_id=port.input_id,
            expected_source_sha256=port.source_sha256,
        )
    except ResolvedNumberingPortError as exc:
        raise ResolvedConversionV4Unsupported(exc.code, str(exc)) from exc
    except ResolvedRuntimeV4Unsupported as exc:
        raise ResolvedConversionV4Unsupported(
            "docwen.resolved_document.invalid",
            str(exc),
        ) from exc
    except RuntimeSemanticsV3Unsupported as exc:
        raise ResolvedConversionV4Unsupported(
            "docwen.resolved_document.invalid",
            str(exc),
        ) from exc
    return PreparedResolvedInputsV4(port, runtime_plan, source_carrier_plan, neutral_path, plan_path)


def compose_resolved_v4_markdown(
    prepared: PreparedResolvedInputsV4,
    resources: ResolvedResourceBinding,
) -> ProjectedResolvedMarkdownV4:
    """Compose markers and private resource paths in original coordinates."""

    source = prepared.port.document.authored_markdown
    runtime_edits = tuple(
        _SourceEdit(
            item.source_start,
            item.source_end,
            item.original,
            item.replacement,
            item.role,
        )
        for item in prepared.runtime_plan.marker_edits
    )
    carrier_edits = tuple(
        _SourceEdit(
            item.source_start,
            item.source_end,
            item.original,
            item.replacement,
            item.role,
        )
        for item in prepared.source_carrier_plan.marker_edits
    )
    resource_edits = tuple(
        _SourceEdit(
            item.source_start,
            item.source_end,
            source[item.source_start : item.source_end],
            item.replacement,
            "resource",
        )
        for item in resources.text_projection.edits
    )

    _prove_projection_identity(source, runtime_edits, prepared.runtime_plan.shielded_source, "marker")
    _prove_projection_identity(
        source,
        carrier_edits,
        prepared.source_carrier_plan.shielded_source,
        "source-carrier",
    )
    _prove_projection_identity(source, resource_edits, resources.rendered_markdown, "resource")
    occurrences = prepared.port.document.resource_occurrences
    if len(resource_edits) != len(occurrences):
        raise ResolvedConversionV4Unsupported(
            "docwen.resolved_document.invalid",
            "resource projection does not enumerate every authenticated occurrence",
        )
    edits = tuple(
        sorted(
            (*runtime_edits, *carrier_edits, *resource_edits),
            key=lambda item: (item.source_start, item.source_end),
        )
    )
    _prove_edit_family(source, edits)
    markdown = _apply_edits(source, edits)
    expected_urls = tuple(resources.path_for(item.resource_id).absolute().as_posix() for item in occurrences)
    return ProjectedResolvedMarkdownV4(
        markdown,
        expected_urls,
        prepared.runtime_plan,
        prepared.source_carrier_plan,
    )


def prove_resolved_v4_image_inventory(
    ast: list[dict[str, Any]],
    projection: ProjectedResolvedMarkdownV4,
) -> None:
    """Require every parsed image to be one authenticated embedded occurrence."""

    actual: list[str] = []
    for node in _walk_ast(ast):
        if node.get("type") != "image":
            continue
        attrs = node.get("attrs")
        url = attrs.get("url") if isinstance(attrs, dict) else None
        if not isinstance(url, str) or not url:
            raise ResolvedConversionV4Unsupported(
                "docwen.resolved_document.invalid",
                "a resolved-v4 image node has no closed URL",
            )
        actual.append(url)
    if Counter(actual) != Counter(projection.expected_image_urls):
        raise ResolvedConversionV4Unsupported(
            "docwen.resolved_document.invalid",
            "parsed image inventory differs from range-bound embedded resources",
        )


def _prove_input_ref(item: Any, *, role: str, kind: str, media_type: str) -> None:
    path = Path(getattr(item, "path", ""))
    if (
        getattr(item, "input_role", None) != role
        or getattr(item, "input_kind", None) != kind
        or getattr(item, "media_type", None) != media_type
        or not path.is_file()
        or path.is_symlink()
    ):
        code = "docwen.numbering_export_plan.invalid" if role == _PLAN_ROLE else "docwen.resolved_document.invalid"
        raise ResolvedConversionV4Unsupported(code, f"{role} input metadata or Workspace copy is invalid")


def _prove_projection_identity(
    source: str,
    edits: tuple[_SourceEdit, ...],
    expected: str,
    role: str,
) -> None:
    _prove_edit_family(source, edits)
    if _apply_edits(source, edits) != expected:
        raise ResolvedConversionV4Unsupported(
            "docwen.resolved_document.invalid",
            f"{role} projection differs from its authenticated source edits",
        )


def _prove_edit_family(source: str, edits: tuple[_SourceEdit, ...]) -> None:
    previous: _SourceEdit | None = None
    for edit in edits:
        if (
            type(edit.source_start) is not int
            or type(edit.source_end) is not int
            or edit.source_start < 0
            or edit.source_end < edit.source_start
            or edit.source_end > len(source)
            or source[edit.source_start : edit.source_end] != edit.original
            or not edit.replacement
        ):
            raise ResolvedConversionV4Unsupported(
                "docwen.resolved_document.invalid",
                "resolved-v4 source edit is not closed or no longer matches authored Markdown",
            )
        if previous is not None:
            overlaps = edit.source_start < previous.source_end
            same_insertion = edit.source_start == previous.source_start and (
                edit.source_start == edit.source_end or previous.source_start == previous.source_end
            )
            if overlaps or same_insertion:
                raise ResolvedConversionV4Unsupported(
                    "docwen.resolved_document.invalid",
                    "resolved-v4 semantic and resource source edits overlap",
                )
        previous = edit


def _apply_edits(source: str, edits: tuple[_SourceEdit, ...]) -> str:
    output = source
    for edit in reversed(edits):
        output = output[: edit.source_start] + edit.replacement + output[edit.source_end :]
    return output


def _walk_ast(nodes: list[dict[str, Any]]):
    for node in nodes:
        yield node
        children = node.get("children")
        if isinstance(children, list):
            yield from _walk_ast(children)
        caption_children = node.get(_CAPTION_CHILDREN_KEY)
        if isinstance(caption_children, list):
            yield from _walk_ast(caption_children)


__all__ = [
    "PreparedResolvedInputsV4",
    "ProjectedResolvedMarkdownV4",
    "ResolvedConversionV4Unsupported",
    "claims_resolved_v4_inputs",
    "compose_resolved_v4_markdown",
    "load_resolved_v4_inputs",
    "prove_resolved_v4_image_inventory",
]
