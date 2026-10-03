"""Production proof-only resolved-v4 DOCX -> Markdown integration gates."""

from __future__ import annotations

import hashlib
import json
from dataclasses import asdict
from pathlib import Path
from typing import Any, cast
from zipfile import ZIP_DEFLATED, ZipFile

import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from lxml import etree

from docwen_core._docx_semantics_v3_model import (
    CaptionStyleBindingV3,
    CaptionStyleKeyV3,
    DocxSemanticsV3Error,
    derive_target_identity_v3,
)
from docwen_core.docx_numbering_import import AMBIGUOUS_VISIBLE_PREFIX_DIAGNOSTIC
from docwen_core.docx_resolved_numbering import ResolvedNumberingDocxSession
from docwen_core.docx_resolved_numbering_recovery import (
    ResolvedNumberingV4Recovery,
)
from docwen_core.markdown_extensions import MarkdownExtensions
from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_core.models.resolved_numbering import (
    NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
    RESOLVED_DOCUMENT_MEDIA_TYPE,
    CaptionMaterialization,
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
    canonicalize_numbering_plan,
)
from docwen_plugin_document.to_markdown.converter import DocxToMarkdownConverter
from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter
from docwen_runtime.config.document_styles import build_document_style_catalog
from docwen_runtime.output.finalizer import OutputFinalizer
from tests.support.cancellation import FakeCancellationTokenView
from tests.support.config import FakeConfigView
from tests.support.execution import FakeExecutionContext
from tests.support.logging import FakePluginLogger
from tests.support.progress import FakeProgressSink
from tests.support.workspace import FakeWorkspaceHandle

pytestmark = pytest.mark.integration

_PROJECT_ROOT = Path(__file__).resolve().parents[2]
_FORWARD_FIXTURES = _PROJECT_ROOT / "packages" / "plugins" / "markdown" / "tests" / "fixtures" / "resolved_v4"
_NEUTRAL = _FORWARD_FIXTURES / "resolved-document.rich.json"
_PLAN = _FORWARD_FIXTURES / "numbering-export-plan.rich.json"


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


def _forward_context(
    tmp_path: Path,
    neutral: Path = _NEUTRAL,
    plan: Path = _PLAN,
) -> FakeExecutionContext:
    staging = tmp_path / "forward-staging"
    staging.mkdir()
    refs = (
        FileRef(
            path=str(neutral),
            format="markdown",
            category="document",
            size_bytes=neutral.stat().st_size,
            input_kind="document",
            input_role="neutral_document",
            logical_path="document.json",
            media_type=RESOLVED_DOCUMENT_MEDIA_TYPE,
        ),
        FileRef(
            path=str(plan),
            format="json",
            category="resource",
            size_bytes=plan.stat().st_size,
            input_kind="resource",
            input_role="numbering_export_plan",
            logical_path="numbering-plan.json",
            media_type=NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
        ),
    )
    request = ConversionRequest(
        request_id="resolved-v4-forward-for-reverse",
        input_refs=list(refs),
        target_format="docx",
        options={"locale": "zh_CN", "heading_merge_mode": "never"},
        output_policy=OutputPolicy(),
    )
    workspace = FakeWorkspaceHandle(str(neutral), str(staging), refs)
    styles = build_document_style_catalog(
        {"gui": {"language": {"locale": "zh_CN"}}},
        locales_dir=_PROJECT_ROOT / "i18n" / "locales",
    )
    return FakeExecutionContext(
        request,
        workspace,
        FakeConfigView(),
        FakeProgressSink(),
        FakeCancellationTokenView(),
        FakePluginLogger(),
        document_style_catalog=styles,
    )


def _reverse_context(tmp_path: Path, source: Path, *, request_id: str) -> FakeExecutionContext:
    staging = tmp_path / f"reverse-staging-{request_id}"
    staging.mkdir()
    ref = FileRef(
        path=str(source),
        format="docx",
        category="document",
        size_bytes=source.stat().st_size,
    )
    return FakeExecutionContext(
        ConversionRequest(
            request_id=request_id,
            input_refs=[ref],
            target_format="md",
            options={
                "to_md_keep_images": True,
                "remove_numbering": True,
                "markdown_extensions": {"output": MarkdownExtensions.obsidian().to_dict()},
            },
            output_policy=OutputPolicy(),
        ),
        FakeWorkspaceHandle(str(source), str(staging), (ref,)),
        FakeConfigView(),
        FakeProgressSink(),
        FakeCancellationTokenView(),
        FakePluginLogger(),
    )


def _forward_representative(tmp_path: Path) -> Path:
    result = MdToDocxConverter().convert(_forward_context(tmp_path))
    assert result.success, result.error
    assert len(result.artifacts) == 1
    finalized = OutputFinalizer().finalize(
        "resolved-v4-forward-for-reverse",
        result.artifacts,
        OutputPolicy(output_dir=str(tmp_path / "published"), overwrite_mode="error"),
    )
    assert finalized.success, finalized.error
    primary = next(item for item in finalized.artifacts if item.is_primary)
    assert not Path(f"{primary.staging_path}.docwen").exists()
    return Path(primary.staging_path)


def test_representative_proves_four_kinds_refs_citation_and_preserves_tokens_without_exact_claim(
    tmp_path: Path,
) -> None:
    source = _forward_representative(tmp_path)
    context = _reverse_context(tmp_path, source, request_id="representative")

    result = DocxToMarkdownConverter().convert(context)

    assert result.success, result.error
    primary = Path(next(item.staging_path for item in result.artifacts if item.is_primary))
    markdown = primary.read_text(encoding="utf-8")
    assert "| Score | 95 |" in markdown
    assert "```rust\nfn main() {}\n```" in markdown
    assert not context.progress.diagnostics
    assert "# Architecture ^h-7f3a" in markdown
    assert "Figure: System overview ^system-overview" in markdown
    assert "Table: Results ^results-main" in markdown
    assert "Equation: ^energy-main" in markdown
    assert "Code: Entry point ^entry-main" in markdown
    assert "Figure: 1" not in markdown
    assert "Table: 1" not in markdown
    assert "Equation: 1" not in markdown
    assert "Code: 1" not in markdown
    assert "Stable: @[[#^h-7f3a]] and @[[#^system-overview|System overview]]." in markdown
    assert "Ordinary: [[#^system-overview]] and ![[Guide#^h-7f3a]]." in markdown
    assert "Citation: @cite-one." in markdown

    reopened = Document(str(source))
    recovery = ResolvedNumberingV4Recovery.load_if_present(source, reopened)
    assert recovery is not None
    assert recovery.caption_signatures == (
        ("figure", "system-overview", "System overview", "1"),
        ("table", "results-main", "Results", "1"),
        ("equation", "energy-main", "", "1"),
        ("code_block", "entry-main", "Entry point", "1"),
    )


def test_bound_caption_cannot_be_downgraded_to_standalone_by_removing_its_carrier(tmp_path: Path) -> None:
    source = _forward_representative(tmp_path)
    tampered = tmp_path / "bound-caption-without-carrier.docx"
    tampered.write_bytes(source.read_bytes())

    with ZipFile(tampered) as package:
        infos = package.infolist()
        members = {item.filename: package.read(item.filename) for item in infos}
    root = etree.fromstring(members["word/document.xml"])
    target_tag = derive_target_identity_v3("figure", "system-overview").tag
    target = next(
        item
        for item in root.iter(qn("w:sdt"))
        if (tag := item.find(f"{qn('w:sdtPr')}/{qn('w:tag')}")) is not None
        and tag.get(qn("w:val")) == target_tag
    )
    content = target.find(qn("w:sdtContent"))
    assert content is not None and len(content) == 2
    content.remove(content[0])
    members["word/document.xml"] = etree.tostring(
        root,
        encoding="UTF-8",
        xml_declaration=True,
        standalone=True,
    )
    rewritten = tampered.with_suffix(".rewrite.docx")
    with ZipFile(rewritten, "w", compression=ZIP_DEFLATED) as output:
        for info in infos:
            output.writestr(info, members[info.filename])
    rewritten.replace(tampered)

    with pytest.raises(DocxSemanticsV3Error, match="caption target SDT"):
        ResolvedNumberingV4Recovery.load_if_present(tampered, Document(str(tampered)))


def test_standalone_captions_round_trip_with_numbering_and_reference_authority(tmp_path: Path) -> None:
    neutral, plan = _write_standalone_caption_pair(tmp_path)
    forward = MdToDocxConverter().convert(_forward_context(tmp_path, neutral, plan))

    assert forward.success, forward.error
    source = Path(forward.artifacts[0].staging_path)
    recovery = ResolvedNumberingV4Recovery.load_if_present(source, Document(str(source)))
    assert recovery is not None
    assert recovery.caption_signatures == (
        ("figure", "fig-plan", "Planned figure", "1"),
        ("table", None, "Planned table", "1"),
        ("code_block", None, "Planned code", ""),
    )
    assert all(not item.object_elements for item in recovery.recovered_captions)

    with ZipFile(source) as package:
        document_xml = package.read("word/document.xml")
        all_xml = b"".join(package.read(name) for name in package.namelist() if name.endswith(".xml"))
    assert document_xml.count(b"docwen-standalone-caption-v1:") == 2
    assert b"https://docwen.dev/schema/document-standalone-caption-occurrence-map/v1" in all_xml
    assert b"SEQ Figure" in document_xml
    assert b"SEQ Table" in document_xml
    assert b"SEQ Code" not in document_xml
    assert b" REF " in document_xml

    reverse = DocxToMarkdownConverter().convert(_reverse_context(tmp_path, source, request_id="standalone-captions"))

    assert reverse.success, reverse.error
    markdown = Path(reverse.artifacts[0].staging_path).read_text(encoding="utf-8")
    assert "Figure: Planned figure ^fig-plan" in markdown
    assert "Table: Planned table" in markdown
    assert "Code: Planned code" in markdown
    assert "See @[[#^fig-plan]]." in markdown
    assert "Figure: 1" not in markdown
    assert "Table: 1" not in markdown


@pytest.mark.parametrize("enabled", [True, False])
def test_manual_heading_prefix_survives_enabled_and_disabled_without_legacy_cleanup(
    tmp_path: Path,
    enabled: bool,
) -> None:
    source = _render_manual_heading_package(tmp_path, enabled=enabled)
    context = _reverse_context(tmp_path, source, request_id=f"manual-{enabled}")

    result = DocxToMarkdownConverter().convert(context)

    assert result.success, result.error
    markdown = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
    assert "# 2.3 标题 ^head-a" in markdown
    assert "# 标题 ^head-a" not in markdown
    assert "# 1 2.3 标题" not in markdown
    diagnostics = [item[2] for item in context.progress.diagnostics]
    assert (AMBIGUOUS_VISIBLE_PREFIX_DIAGNOSTIC in diagnostics) is (not enabled)


@pytest.mark.parametrize("conflicting_paragraph_style", [False, True])
def test_office_renumbered_caption_styles_keep_authenticated_bindings(
    tmp_path: Path, conflicting_paragraph_style: bool
) -> None:
    source = _forward_representative(tmp_path)
    original = ResolvedNumberingV4Recovery.load_if_present(source, Document(str(source)))
    assert original is not None
    replacements = {
        binding.resolved_style_id: f"OfficeCaption{index}"
        for index, binding in enumerate(original.caption_style_bindings)
    }
    with ZipFile(source) as package:
        infos = package.infolist()
        members = {item.filename: package.read(item.filename) for item in infos}
    for member in ("word/styles.xml", "word/document.xml"):
        root = etree.fromstring(members[member])
        for element in root.iter():
            attribute = qn("w:styleId") if element.tag == qn("w:style") else qn("w:val")
            if element.tag in {qn(f"w:{tag}") for tag in ("style", "pStyle", "basedOn", "next", "link")}:
                current = element.get(attribute)
                if current in replacements:
                    element.set(attribute, replacements[current])
        if member == "word/document.xml" and conflicting_paragraph_style:
            table_style = replacements[original.caption_style_bindings[1].resolved_style_id]
            figure_style = replacements[original.caption_style_bindings[0].resolved_style_id]
            paragraph_style = next(item for item in root.iter(qn("w:pStyle")) if item.get(qn("w:val")) == table_style)
            paragraph_style.set(qn("w:val"), figure_style)
        members[member] = etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
    saved = tmp_path / "office-saved.docx"
    with ZipFile(saved, "w", compression=ZIP_DEFLATED) as package:
        for info in infos:
            package.writestr(info, members[info.filename])
    context = _reverse_context(tmp_path, saved, request_id="office-styles")

    result = DocxToMarkdownConverter().convert(context)

    if conflicting_paragraph_style:
        assert not result.success
        assert result.artifacts == []
        assert context.workspace.registered_artifacts == []
        assert list(Path(context.workspace.staging_dir).iterdir()) == []
        return
    assert result.success, result.error
    recovered = ResolvedNumberingV4Recovery.load_if_present(saved, Document(str(saved)))
    assert recovered is not None
    assert recovered.caption_signatures == original.caption_signatures
    markdown = Path(next(item.staging_path for item in result.artifacts if item.is_primary)).read_text(encoding="utf-8")
    for expected in (
        "Figure: System overview ^system-overview",
        "Table: Results ^results-main",
        "Equation: ^energy-main",
        "Code: Entry point ^entry-main",
        "@[[#^system-overview|System overview]]",
        "Citation: @cite-one.",
    ):
        assert expected in markdown


def test_reference_cache_tamper_fails_before_artifact_or_staging_publish(tmp_path: Path) -> None:
    source = _forward_representative(tmp_path)
    tampered = tmp_path / "tampered-reference.docx"
    tampered.write_bytes(source.read_bytes())
    _tamper_first_reference_cached_result(tampered)
    context = _reverse_context(tmp_path, tampered, request_id="tampered")

    result = DocxToMarkdownConverter().convert(context)

    assert not result.success
    assert result.artifacts == []
    assert context.workspace.registered_artifacts == []
    assert list(Path(context.workspace.staging_dir).iterdir()) == []


@pytest.mark.parametrize("legacy_companion", [False, True])
def test_docx_alone_reconstructs_semantics_and_ignores_unrelated_companions(
    tmp_path: Path,
    legacy_companion: bool,
) -> None:
    generated = _forward_representative(tmp_path)
    isolated = tmp_path / "isolated-input"
    isolated.mkdir()
    source = isolated / "document.docx"
    source.write_bytes(generated.read_bytes())
    if legacy_companion:
        source.with_suffix(".docx.docwen").write_bytes(b"unrelated-file-must-not-be-read")
    context = _reverse_context(tmp_path, source, request_id="isolated")

    result = DocxToMarkdownConverter().convert(context)

    assert result.success, result.error
    markdown = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
    assert "# Architecture ^h-7f3a" in markdown
    assert "Table: Results ^results-main" in markdown
    assert "| Score | 95 |" in markdown
    assert "@[[#^system-overview|System overview]]" in markdown
    assert "@cite-one" in markdown
    assert not context.progress.diagnostics
    if legacy_companion:
        assert source.with_suffix(".docx.docwen").read_bytes() == b"unrelated-file-must-not-be-read"


def test_document_edit_is_recovered_from_current_docx(tmp_path: Path) -> None:
    source = _forward_representative(tmp_path)
    _replace_zip_member_bytes(source, member="word/document.xml", old=b"<w:t>95</w:t>", new=b"<w:t>96</w:t>")
    context = _reverse_context(tmp_path, source, request_id="edited-docx")

    result = DocxToMarkdownConverter().convert(context)

    assert result.success, result.error
    markdown = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
    expected_source = json.loads(_NEUTRAL.read_text(encoding="utf-8"))["document"]["authored_markdown"]
    assert markdown != expected_source
    assert "| Score | 96 |" in markdown
    assert "| Score | 95 |" not in markdown
    assert not context.progress.diagnostics


def _caption_bindings(document: Any) -> tuple[CaptionStyleBindingV3, ...]:
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


def _manual_port(*, enabled: bool) -> ResolvedNumberingPort:
    source = "# 2.3 标题 ^head-a\n\nTable: 数据\n\n| A |\n|---|\n| B |\n"
    heading_end = source.index("\n")
    table_start = source.index("Table:")
    table_end = len(source) - 1
    heading = ResolvedDocumentTarget(0, heading_end, _sha(source[:heading_end]), "heading", "head-a", 1, "2.3 标题")
    table = ResolvedDocumentTarget(
        table_start,
        table_end,
        _sha(source[table_start:table_end]),
        "table",
        None,
        None,
        "数据",
    )
    definition = HeadingDefinition(
        "main",
        (
            HeadingLevelDefinition(
                1,
                1,
                "arabic_half",
                (HeadingCounterSegment(1, "arabic_half"),),
                "space",
                None,
            ),
        ),
    )
    heading_plan = NumberingTarget(
        0,
        heading_end,
        "heading",
        enabled,
        "head-a",
        "1" if enabled else None,
        HeadingListMaterialization("main", "document", 1) if enabled else None,
    )
    plan = ResolvedNumberingPlan(
        (definition,) if enabled else (),
        (HeadingInstance("document", "main", ()),) if enabled else (),
        (
            heading_plan,
            NumberingTarget(table_start, table_end, "table", False, None, None, None),
        ),
    )
    identity = "a" * 64
    document = ResolvedDocument(source, (heading, table), (), (), (), ())
    return ResolvedNumberingPort(
        ResolvedDocumentEnvelope("manual", _sha(source), identity, document),
        NumberingExportPlanEnvelope("manual", _sha(source), identity, plan),
    )


def _render_manual_heading_package(tmp_path: Path, *, enabled: bool) -> Path:
    document = Document()
    session = ResolvedNumberingDocxSession(
        document,
        _manual_port(enabled=enabled),
        heading_style_ids={1: "Heading1"},
        heading_style_names={"heading_1": "Heading 1"},
        caption_style_bindings=_caption_bindings(document),
    )
    heading = document.add_heading("2.3 标题", level=1)
    caption = document.add_paragraph(style="Table Caption V4")
    table = document.add_table(rows=2, cols=1)
    table.cell(0, 0).text = "A"
    table.cell(1, 0).text = "B"
    targets = session.port.document.targets
    session.bind_heading(heading, source_start=targets[0].source_start, source_end=targets[0].source_end)
    session.bind_caption(
        caption,
        (table._element,),
        source_start=targets[1].source_start,
        source_end=targets[1].source_end,
        kind="table",
    )
    output = tmp_path / f"manual-{enabled}.docx"
    session.write_package(output)
    return output


def _tamper_first_reference_cached_result(path: Path) -> None:
    with ZipFile(path) as package:
        infos = package.infolist()
        members = {item.filename: package.read(item.filename) for item in infos}
    root = etree.fromstring(members["word/document.xml"])
    occurrence = next(
        item
        for item in root.iter(qn("w:sdt"))
        if (tag := item.find(f"{qn('w:sdtPr')}/{qn('w:tag')}")) is not None
        and (tag.get(qn("w:val")) or "").startswith("docwen-ref-occurrence-v1:")
    )
    result_texts = list(occurrence.iter(qn("w:t")))
    assert result_texts and result_texts[0].text == "1"
    result_texts[0].text = "9"
    members["word/document.xml"] = etree.tostring(
        root,
        encoding="UTF-8",
        xml_declaration=True,
        standalone=True,
    )
    rewritten = path.with_suffix(".rewrite.docx")
    with ZipFile(rewritten, "w", compression=ZIP_DEFLATED) as output:
        for info in infos:
            output.writestr(info, members[info.filename])
    rewritten.replace(path)


def _replace_zip_member_bytes(path: Path, *, member: str, old: bytes, new: bytes) -> None:
    with ZipFile(path) as package:
        infos = package.infolist()
        members = {item.filename: package.read(item.filename) for item in infos}
    assert members[member].count(old) == 1
    members[member] = members[member].replace(old, new)
    rewritten = path.with_suffix(".rewrite.docx")
    with ZipFile(rewritten, "w", compression=ZIP_DEFLATED) as output:
        for info in infos:
            output.writestr(info, members[info.filename])
    rewritten.replace(path)
