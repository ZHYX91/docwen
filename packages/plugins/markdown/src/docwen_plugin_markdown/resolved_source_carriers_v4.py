"""Profile-free source-carrier bridge for the resolved v4 conversion port.

The resolved-document and numbering-plan inputs remain the sole authority for
targets, numbering, cross-references, and Citations.  This module consults the
frozen Markdown source oracle only for the two lossless source carriers that
are not duplicated in the resolved wire model:

* ordinary block anchors and their source-container topology; and
* exact fenced-source framing.

The v3 oracle still projects legacy numbering fields for compatibility.  They
are deliberately neither copied nor inspected here.  Likewise its target,
reference, and Citation runtime markers are discarded before Markdown is
parsed, so they cannot become a second numbering or resolver authority.
"""

from __future__ import annotations

from dataclasses import dataclass

from docwen_plugin_markdown.document_semantics_v3_fenced_source import (
    fenced_source_info_insertion_offset_v3,
)
from docwen_plugin_markdown.runtime_semantics_v3 import (
    RuntimeMarkerV3,
    RuntimeSemanticsV3Plan,
    RuntimeSemanticsV3Unsupported,
    apply_runtime_semantics_v3,
    prepare_runtime_semantics_v3,
)


@dataclass(frozen=True, slots=True)
class ResolvedSourceCarrierEditV4:
    """One source-authenticated carrier marker edit."""

    source_start: int
    source_end: int
    original: str
    replacement: str
    role: str


@dataclass(frozen=True, slots=True)
class ResolvedSourceCarrierPlanV4:
    """Carrier-only projection derived from one immutable authored source."""

    runtime_plan: RuntimeSemanticsV3Plan
    marker_edits: tuple[ResolvedSourceCarrierEditV4, ...]

    @property
    def source_sha256(self) -> str:
        return self.runtime_plan.source_sha256

    @property
    def shielded_source(self) -> str:
        return self.runtime_plan.shielded_source


def prepare_resolved_source_carriers_v4(
    source: str,
    *,
    input_id: str,
    expected_source_sha256: str,
) -> ResolvedSourceCarrierPlanV4:
    """Project only anchors and fences from the frozen source oracle.

    A source diagnostic remains fatal at this boundary.  The resolved port may
    not use a provider-authored plan to bypass invalid authored Markdown.
    """

    full_plan = prepare_runtime_semantics_v3(source, input_id=input_id)
    if full_plan.source_sha256 != expected_source_sha256:
        raise RuntimeSemanticsV3Unsupported("source-carrier projection belongs to a different authenticated source")
    if full_plan.analysis.has_errors:
        codes = ", ".join(sorted({str(item["code"]) for item in full_plan.analysis.diagnostics}))
        raise RuntimeSemanticsV3Unsupported(f"authored Markdown has source-oracle diagnostics: {codes or 'unknown'}")

    carrier_markers = tuple(
        marker for marker in full_plan.markers if marker.role in {"ordinary_anchor", "fenced_source"}
    )
    edits = tuple(_carrier_edit(source, marker) for marker in carrier_markers)
    _prove_edit_family(source, edits)
    shielded_source = _apply_edits(source, edits)
    carrier_plan = RuntimeSemanticsV3Plan(
        analysis=full_plan.analysis,
        shielded_source=shielded_source,
        markers=carrier_markers,
        body_start=full_plan.body_start,
        ordinary_anchor_parents=full_plan.ordinary_anchor_parents,
    )
    return ResolvedSourceCarrierPlanV4(runtime_plan=carrier_plan, marker_edits=edits)


def apply_resolved_source_carriers_v4(
    ast_nodes: list[dict[str, object]],
    plan: ResolvedSourceCarrierPlanV4,
) -> list[dict[str, object]]:
    """Restore and structurally bind every carrier marker exactly once."""

    # RuntimeSemanticsV3 works on the same JSON-like AST shape.  Keep the cast
    # local to this profile-free adapter instead of broadening either public
    # runtime contract.
    return apply_runtime_semantics_v3(ast_nodes, plan.runtime_plan)  # type: ignore[arg-type,return-value]


def _carrier_edit(source: str, marker: RuntimeMarkerV3) -> ResolvedSourceCarrierEditV4:
    if marker.role == "ordinary_anchor":
        source_range = marker.payload.get("range")
        if not isinstance(source_range, dict):
            raise RuntimeSemanticsV3Unsupported("ordinary-anchor marker has no source range")
        start = source_range.get("start")
        end = source_range.get("end")
        if type(start) is not int or type(end) is not int:
            raise RuntimeSemanticsV3Unsupported("ordinary-anchor source range is not integral")
        return ResolvedSourceCarrierEditV4(
            source_start=start,
            source_end=end,
            original=source[start:end],
            replacement=marker.marker,
            role=marker.role,
        )
    if marker.role == "fenced_source":
        record = marker.payload.get("record")
        if not isinstance(record, dict):
            raise RuntimeSemanticsV3Unsupported("fenced-source marker has no canonical record")
        offset = fenced_source_info_insertion_offset_v3(source, record)
        return ResolvedSourceCarrierEditV4(
            source_start=offset,
            source_end=offset,
            original="",
            replacement=f" {marker.marker}",
            role=marker.role,
        )
    raise RuntimeSemanticsV3Unsupported("resolved source-carrier plan contains a forbidden marker role")


def _prove_edit_family(source: str, edits: tuple[ResolvedSourceCarrierEditV4, ...]) -> None:
    previous_end = -1
    insertion_points: set[int] = set()
    for edit in edits:
        if (
            edit.source_start < 0
            or edit.source_end < edit.source_start
            or edit.source_end > len(source)
            or source[edit.source_start : edit.source_end] != edit.original
            or not edit.replacement
        ):
            raise RuntimeSemanticsV3Unsupported("source-carrier edit is outside the authored source")
        if edit.source_start < previous_end:
            raise RuntimeSemanticsV3Unsupported("source-carrier edits overlap")
        if edit.source_start == edit.source_end:
            if edit.source_start in insertion_points:
                raise RuntimeSemanticsV3Unsupported("source-carrier insertion points collide")
            insertion_points.add(edit.source_start)
        else:
            previous_end = edit.source_end


def _apply_edits(source: str, edits: tuple[ResolvedSourceCarrierEditV4, ...]) -> str:
    output = source
    for edit in reversed(edits):
        output = output[: edit.source_start] + edit.replacement + output[edit.source_end :]
    return output


__all__ = [
    "ResolvedSourceCarrierEditV4",
    "ResolvedSourceCarrierPlanV4",
    "apply_resolved_source_carriers_v4",
    "prepare_resolved_source_carriers_v4",
]
