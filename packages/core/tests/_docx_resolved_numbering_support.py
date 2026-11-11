"""Shared resolved-numbering fixtures and package mutation helpers."""

from __future__ import annotations

import hashlib
from collections.abc import Callable
from pathlib import Path
from typing import cast
from zipfile import ZipFile

from docx.enum.style import WD_STYLE_TYPE

from docwen_core._docx_semantics_v3_model import CaptionStyleBindingV3, CaptionStyleKeyV3
from docwen_core.models.resolved_numbering import (
    HeadingCounterSegment,
    HeadingDefinition,
    HeadingInstance,
    HeadingLevelDefinition,
    HeadingListMaterialization,
    NumberingExportPlanEnvelope,
    NumberingTarget,
    ResolvedDocument,
    ResolvedDocumentEnvelope,
    ResolvedDocumentTarget,
    ResolvedNumberingPlan,
    ResolvedNumberingPort,
    ResolvedReference,
)


def _sha(value: str) -> str:
    return hashlib.sha256(value.encode()).hexdigest()


def _caption_bindings(document) -> tuple[CaptionStyleBindingV3, ...]:
    output: list[CaptionStyleBindingV3] = []
    for semantic_key, style_id, name in (
        ("figure_caption", "DWFigureCaption", "Figure Caption V4"),
        ("table_caption", "DWTableCaption", "Table Caption V4"),
        ("equation_caption", "DWEquationCaption", "Equation Caption V4"),
        ("code_block_caption", "DWCodeCaption", "Code Caption V4"),
    ):
        style = document.styles.add_style(style_id, WD_STYLE_TYPE.PARAGRAPH)
        style.name = name
        output.append(CaptionStyleBindingV3(cast(CaptionStyleKeyV3, semantic_key), style.style_id, style.name))
    return tuple(output)


def _port() -> ResolvedNumberingPort:
    source = "# 2.3 标题\n\nTable: 数据\n|a|\n|-|\n|b|\n\n@[[#^head-a|标题]]\n"
    heading_start = 0
    heading_end = source.index("\n")
    table_start = source.index("Table:")
    table_end = source.index("\n\n@[[")
    token = "@[[#^head-a|标题]]"
    reference_start = source.index(token)
    reference_end = reference_start + len(token)
    heading_target = ResolvedDocumentTarget(
        source_start=heading_start,
        source_end=heading_end,
        source_slice_sha256=_sha(source[heading_start:heading_end]),
        kind="heading",
        target_id="head-a",
        heading_level=1,
        authored_text="2.3 标题",
    )
    table_target = ResolvedDocumentTarget(
        source_start=table_start,
        source_end=table_end,
        source_slice_sha256=_sha(source[table_start:table_end]),
        kind="table",
        target_id=None,
        heading_level=None,
        authored_text="数据",
    )
    reference = ResolvedReference(
        source_start=reference_start,
        source_end=reference_end,
        source_slice_sha256=_sha(token),
        authored_token=token,
        target_source_start=heading_start,
        target_source_end=heading_end,
        target_kind="heading",
        target_id="head-a",
        cached_number="1",
        alias="标题",
    )
    definition = HeadingDefinition(
        definition_id="main",
        levels=(
            HeadingLevelDefinition(
                level=1,
                start=1,
                number_format="arabic_half",
                display=(HeadingCounterSegment(1, "arabic_half"),),
                suffix="space",
                restart_after_level=None,
            ),
        ),
    )
    plan = ResolvedNumberingPlan(
        heading_definitions=(definition,),
        heading_instances=(HeadingInstance("document", "main", ()),),
        targets=(
            NumberingTarget(
                source_start=heading_start,
                source_end=heading_end,
                kind="heading",
                enabled=True,
                target_id="head-a",
                derived_number="1",
                materialization=HeadingListMaterialization("main", "document", 1),
            ),
            NumberingTarget(
                source_start=table_start,
                source_end=table_end,
                kind="table",
                enabled=False,
                target_id=None,
                derived_number=None,
                materialization=None,
            ),
        ),
    )
    plan_sha = "a" * 64
    return ResolvedNumberingPort(
        ResolvedDocumentEnvelope(
            input_id="source",
            source_sha256=_sha(source),
            plan_sha256=plan_sha,
            document=ResolvedDocument(
                source,
                (heading_target, table_target),
                (reference,),
                (),
                (),
                (),
            ),
        ),
        NumberingExportPlanEnvelope(
            input_id="source",
            source_sha256=_sha(source),
            plan_sha256=plan_sha,
            plan=plan,
        ),
    )


def _mutate_zip_member(path: Path, member: str, mutate: Callable[[bytes], bytes]) -> None:
    with ZipFile(path) as package:
        infos = package.infolist()
        data = {info.filename: package.read(info.filename) for info in infos}
    data[member] = mutate(data[member])
    temporary = path.with_suffix(".rewrite.docx")
    with ZipFile(temporary, "w") as output:
        for info in infos:
            output.writestr(info, data[info.filename])
    temporary.replace(path)
