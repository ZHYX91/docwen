"""Source-relative projection for resolved inline DOCX carriers.

The caller has already rendered the paragraph and bound each resolved
reference/Citation to one direct inline SDT.  This module proves that the SDTs
occupy their authenticated positions inside the target's exact authored title,
then captures an immutable OOXML projection for package reopen.
"""

from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass
from typing import Any

from docwen_core._docx_semantics_v3_ooxml import sdt_tag, visible_text

_INLINE_CARRIER_PREFIXES = (
    "docwen-soft-ref-v1:",
    "docwen-ref-occurrence-v1:",
    "docwen-citation-occurrence-v1:",
)


class ResolvedInlineProjectionError(ValueError):
    """The rendered inline payload differs from its authenticated source."""


@dataclass(frozen=True, slots=True)
class ResolvedInlineCarrier:
    source_start: int
    source_end: int
    tag: str
    authored_token: str


@dataclass(frozen=True, slots=True)
class ResolvedInlineFragment:
    source_start: int
    source_end: int
    carrier_tag: str | None
    elements: tuple[Any, ...]


def capture_resolved_inline_projection(
    paragraph_element: Any,
    *,
    authored_source: str,
    target_start: int,
    target_end: int,
    authored_title: str,
    carriers: tuple[ResolvedInlineCarrier, ...],
    allow_rendered_suffix: bool,
) -> tuple[ResolvedInlineFragment, ...]:
    """Prove and snapshot one exact title projection.

    ``allow_rendered_suffix`` supports the Heading merge contract: the exact
    authored title must remain the paragraph prefix, while a caller-rendered
    suffix may follow it.  Captions have no such suffix and use ``False``.
    """

    from docx.oxml.ns import qn

    title_start, title_end = _unique_title_span(
        authored_source,
        target_start=target_start,
        target_end=target_end,
        authored_title=authored_title,
    )
    ordered = tuple(sorted(carriers, key=lambda item: (item.source_start, item.source_end, item.tag)))
    if len({item.tag for item in ordered}) != len(ordered):
        raise ResolvedInlineProjectionError("resolved inline carrier tag is duplicated")
    for carrier in ordered:
        if not title_start <= carrier.source_start < carrier.source_end <= title_end:
            raise ResolvedInlineProjectionError(
                "resolved inline carrier is outside the target's exact authored title span"
            )
        if authored_source[carrier.source_start : carrier.source_end] != carrier.authored_token:
            raise ResolvedInlineProjectionError("resolved inline carrier token differs from its source range")

    payload = tuple(item for item in paragraph_element if item.tag != qn("w:pPr"))
    direct_tags = tuple(
        tag
        for item in payload
        if item.tag == qn("w:sdt") and (tag := sdt_tag(item)) is not None and tag.startswith(_INLINE_CARRIER_PREFIXES)
    )
    nested_tags = tuple(
        value
        for item in payload
        for tag_element in item.iter(qn("w:tag"))
        if (value := tag_element.get(qn("w:val"))) is not None and value.startswith(_INLINE_CARRIER_PREFIXES)
    )
    expected_tags = tuple(item.tag for item in ordered)
    if direct_tags != expected_tags or nested_tags != expected_tags:
        raise ResolvedInlineProjectionError(
            "resolved inline carrier topology/order differs from its source-relative projection"
        )

    by_tag = {item.tag: item for item in ordered}
    fragments: list[ResolvedInlineFragment] = []
    plain: list[Any] = []
    cursor = title_start
    for element in payload:
        tag = sdt_tag(element) if element.tag == qn("w:sdt") else None
        carrier = by_tag.get(tag) if tag is not None else None
        if carrier is None:
            plain.append(element)
            continue
        _append_plain_fragment(
            fragments,
            authored_source,
            cursor,
            carrier.source_start,
            plain,
            allow_suffix=False,
        )
        fragments.append(
            ResolvedInlineFragment(
                carrier.source_start,
                carrier.source_end,
                carrier.tag,
                (deepcopy(element),),
            )
        )
        plain.clear()
        cursor = carrier.source_end
    _append_plain_fragment(
        fragments,
        authored_source,
        cursor,
        title_end,
        plain,
        allow_suffix=allow_rendered_suffix,
    )
    return tuple(fragments)


def _unique_title_span(
    authored_source: str,
    *,
    target_start: int,
    target_end: int,
    authored_title: str,
) -> tuple[int, int]:
    target_slice = authored_source[target_start:target_end]
    relative_start = target_slice.find(authored_title)
    if relative_start < 0 or target_slice.find(authored_title, relative_start + 1) >= 0:
        raise ResolvedInlineProjectionError("target authored title is not unique inside its authenticated source range")
    start = target_start + relative_start
    return start, start + len(authored_title)


def _append_plain_fragment(
    output: list[ResolvedInlineFragment],
    authored_source: str,
    source_start: int,
    source_end: int,
    elements: list[Any],
    *,
    allow_suffix: bool,
) -> None:
    expected = authored_source[source_start:source_end]
    rendered = "".join(visible_text(item) for item in elements)
    valid = rendered.startswith(expected) if allow_suffix else rendered == expected
    if not valid:
        raise ResolvedInlineProjectionError(
            "rendered inline plain fragment differs from its exact source-relative projection"
        )
    output.append(
        ResolvedInlineFragment(
            source_start,
            source_end,
            None,
            tuple(deepcopy(item) for item in elements),
        )
    )


__all__ = [
    "ResolvedInlineCarrier",
    "ResolvedInlineFragment",
    "ResolvedInlineProjectionError",
    "capture_resolved_inline_projection",
]
