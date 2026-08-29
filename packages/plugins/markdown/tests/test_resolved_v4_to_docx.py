"""Production exact-two converter/renderer/session integration gates."""

from __future__ import annotations

import hashlib
import json
import shutil
from dataclasses import asdict
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import pytest
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree
from tests.support.cancellation import FakeCancellationTokenView
from tests.support.config import FakeConfigView
from tests.support.execution import FakeExecutionContext
from tests.support.logging import FakePluginLogger
from tests.support.progress import FakeProgressSink
from tests.support.workspace import FakeWorkspaceHandle

from docwen_core.docx_resolved_numbering import (
    ResolvedNumberingDocxError,
    ResolvedNumberingDocxSession,
)
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
    NumberingTarget,
    ResolvedCitation,
    ResolvedCitationItem,
    ResolvedDocument,
    ResolvedDocumentTarget,
    ResolvedNumberingPlan,
    ResolvedReference,
    canonicalize_numbering_plan,
)
from docwen_core.round_trip_sidecar import ROUND_TRIP_SIDECAR_MEDIA_TYPE, read_round_trip_sidecar
from docwen_plugin_markdown import renderer as markdown_renderer
from docwen_plugin_markdown.manifest import RESOLVED_V4_MD_TO_DOCX_OPTIONS_SCHEMA
from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter
from docwen_runtime.config.document_styles import build_document_style_catalog

from .conftest import PROJECT_ROOT

pytestmark = pytest.mark.contract

_FIXTURES = Path(__file__).parent / "fixtures" / "resolved_v4"
_NEUTRAL = _FIXTURES / "resolved-document.rich.json"
_PLAN = _FIXTURES / "numbering-export-plan.rich.json"


def test_equation_snapshot_cleanup_preserves_whitespace_only_leaf_text() -> None:
    equation = OxmlElement("m:oMath")
    equation.text = "\n  "
    run = OxmlElement("m:r")
    run.text = "\n    "
    leaf = OxmlElement("m:t")
    leaf.set(qn("xml:space"), "preserve")
    leaf.text = " "
    leaf.tail = "\n  "
    run.append(leaf)
    equation.append(run)

    markdown_renderer._strip_serialization_only_whitespace(equation)  # pyright: ignore[reportPrivateUsage]

    assert equation.text is None
    assert run.text is None
    assert leaf.text == " "
    assert leaf.tail is None


def test_active_resolved_v4_options_are_exactly_three() -> None:
    import docwen_application.conversion_service as application_conversion_service

    application_schema = application_conversion_service._MARKDOWN_TO_DOCX_OPTIONS  # pyright: ignore[reportPrivateUsage]
    assert set(RESOLVED_V4_MD_TO_DOCX_OPTIONS_SCHEMA["properties"]) == {
        "locale",
        "template_name",
        "heading_merge_mode",
    }
    assert RESOLVED_V4_MD_TO_DOCX_OPTIONS_SCHEMA["additionalProperties"] is False
    assert application_schema == RESOLVED_V4_MD_TO_DOCX_OPTIONS_SCHEMA


def _sha_text(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def _refs(neutral: Path, plan: Path) -> tuple[FileRef, FileRef]:
    return (
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


def _context(
    tmp_path: Path,
    refs: tuple[FileRef, ...],
    *,
    options: dict[str, Any] | None = None,
    style_locale: str = "zh_CN",
) -> tuple[FakeExecutionContext, FakeWorkspaceHandle]:
    staging = tmp_path / "staging"
    staging.mkdir(parents=True)
    request = ConversionRequest(
        request_id="resolved-v4-test",
        input_refs=list(refs),
        target_format="docx",
        options=options or {"locale": "zh_CN", "heading_merge_mode": "never"},
        output_policy=OutputPolicy(),
    )
    workspace = FakeWorkspaceHandle(refs[0].path if refs else "", str(staging), refs)
    styles = build_document_style_catalog(
        {"gui": {"language": {"locale": style_locale}}},
        locales_dir=PROJECT_ROOT / "i18n" / "locales",
    )
    return (
        FakeExecutionContext(
            request,
            workspace,
            FakeConfigView(),
            FakeProgressSink(),
            FakeCancellationTokenView(),
            FakePluginLogger(),
            document_style_catalog=styles,
        ),
        workspace,
    )


def _forbidden_legacy(*_args: Any, **_kwargs: Any) -> Any:
    raise AssertionError("resolved-v4 request entered a historical source converter")


def test_representative_exact_two_materializes_all_physical_semantics_without_legacy(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_plugin_markdown.to_docx.converter as converter_module

    for name in (
        "load_bibliography_resource",
        "read_input_markdown",
        "prepare_runtime_semantics_v3",
        "remove_md_numbering",
        "add_md_numbering",
        "process_markdown_links",
        "bind_declared_markdown_images",
    ):
        monkeypatch.setattr(converter_module, name, _forbidden_legacy)
    context, workspace = _context(tmp_path, _refs(_NEUTRAL, _PLAN))

    result = MdToDocxConverter().convert(context)

    assert result.success, result.error
    assert [item.code for item in result.diagnostics if item.level == "error"] == []
    assert len(result.artifacts) == len(workspace.registered_artifacts) == 2
    output = Path(next(item.staging_path for item in result.artifacts if item.is_primary))
    sidecar = next(item for item in result.artifacts if item.media_type == ROUND_TRIP_SIDECAR_MEDIA_TYPE)
    assert sorted(item.name for item in Path(workspace.staging_dir).iterdir()) == sorted(
        [output.name, Path(sidecar.staging_path).name]
    )
    recovered = read_round_trip_sidecar(sidecar.staging_path, docx_path=output)
    assert recovered.neutral_document == _NEUTRAL.read_bytes()
    assert recovered.numbering_export_plan == _PLAN.read_bytes()
    assert (
        recovered.authored_source
        == json.loads(_NEUTRAL.read_text(encoding="utf-8"))["document"]["authored_markdown"].encode()
    )
    with ZipFile(output) as package:
        names = package.namelist()
        document_bytes = package.read("word/document.xml")
        document = etree.fromstring(document_bytes)
        instructions = [item.text or "" for item in document.iter(qn("w:instrText"))]
        visible = "".join(item.text or "" for item in document.iter(qn("w:t")))
        all_xml = b"".join(package.read(name) for name in names if name.endswith(".xml"))
    assert "word/numbering.xml" in names
    assert len([item for item in names if item.startswith("word/media/")]) == 1
    assert all(
        any(token in instruction for instruction in instructions)
        for token in ("SEQ Figure", "SEQ Table", "SEQ Equation", "SEQ Code")
    )
    assert sum(" REF " in item for item in instructions) == 2
    assert sum(" CITATION " in item for item in instructions) == 1
    assert "Architecture" in visible
    assert "[[#^system-overview]] and ![[Guide#^h-7f3a]]" in visible
    assert "One, A. (2026). Exact two." in visible
    assert b"resolved-v4-resources" not in all_xml
    assert b"https://docwen.dev/schema/document-fenced-source-map/v1" in all_xml
    assert b"docwen-fenced-source-v1:" in document_bytes


def test_exact_two_preserves_nested_fence_anchor_and_topology_carriers(tmp_path: Path) -> None:
    neutral_payload = json.loads(_NEUTRAL.read_text(encoding="utf-8"))
    plan_payload = json.loads(_PLAN.read_text(encoding="utf-8"))
    suffix = "\n> ```mermaid\n> graph TD\n> ```\n>\n> ^inner-fence\n\n^outer-quote\n\nTop A ^top-a\n"
    source = neutral_payload["document"]["authored_markdown"] + suffix
    source_sha256 = _sha_text(source)
    neutral_payload["source_sha256"] = source_sha256
    neutral_payload["document"]["authored_markdown"] = source
    plan_payload["source_sha256"] = source_sha256
    neutral = tmp_path / "neutral.json"
    plan = tmp_path / "plan.json"
    neutral.write_text(json.dumps(neutral_payload, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    plan.write_text(json.dumps(plan_payload, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    context, workspace = _context(tmp_path, _refs(neutral, plan))

    result = MdToDocxConverter().convert(context)

    assert result.success, result.error
    output = Path(result.artifacts[0].staging_path)
    with ZipFile(output) as package:
        document_bytes = package.read("word/document.xml")
        document = etree.fromstring(document_bytes)
        visible = "".join(item.text or "" for item in document.iter(qn("w:t")))
        all_xml = b"".join(package.read(name) for name in package.namelist() if name.endswith(".xml"))
    assert b"docwen-fenced-source-v1:" in document_bytes
    assert document_bytes.count(b"docwen-anchor-v1:") == 3
    assert b"https://docwen.dev/schema/document-fenced-source-map/v1" in all_xml
    assert b"https://docwen.dev/schema/document-anchor-topology-map/v1" in all_xml
    assert "^inner-fence" not in visible
    assert "^outer-quote" not in visible
    assert "^top-a" not in visible
    assert len(workspace.registered_artifacts) == 2


def test_partial_claim_and_legacy_options_fail_before_source_pipeline(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_plugin_markdown.to_docx.converter as converter_module

    for name in ("read_input_markdown", "prepare_runtime_semantics_v3", "remove_md_numbering", "add_md_numbering"):
        monkeypatch.setattr(converter_module, name, _forbidden_legacy)
    neutral_ref, plan_ref = _refs(_NEUTRAL, _PLAN)

    partial_context, partial_workspace = _context(tmp_path / "partial", (neutral_ref,))
    partial = MdToDocxConverter().convert(partial_context)
    assert not partial.success
    assert partial.error is not None
    assert partial.error.diagnostic_code == "docwen.numbering_export_plan.missing"
    assert partial_workspace.registered_artifacts == []
    assert list(Path(partial_workspace.staging_dir).iterdir()) == []

    legacy_context, legacy_workspace = _context(
        tmp_path / "legacy",
        (neutral_ref, plan_ref),
        options={"remove_numbering": True},
    )
    legacy = MdToDocxConverter().convert(legacy_context)
    assert not legacy.success
    assert legacy.error is not None
    assert legacy.error.diagnostic_code == "MD2DOCX-RESOLVED-V4-OPTIONS-INVALID"
    assert legacy_workspace.registered_artifacts == []
    assert list(Path(legacy_workspace.staging_dir).iterdir()) == []


@pytest.mark.parametrize(
    ("options", "style_locale"),
    [
        ({"locale": True, "heading_merge_mode": "never"}, "zh_CN"),
        ({"locale": "en_US", "heading_merge_mode": "never"}, "zh_CN"),
        ({"locale": "zh_CN", "template_name": 1, "heading_merge_mode": "never"}, "zh_CN"),
        ({"locale": "zh_CN", "template_name": {}, "heading_merge_mode": "never"}, "zh_CN"),
        ({"locale": "zh_CN", "template_name": "  ", "heading_merge_mode": "never"}, "zh_CN"),
        ({"locale": "zh_CN", "heading_merge_mode": False}, "zh_CN"),
        ({"locale": "zh_CN", "heading_merge_mode": 1}, "zh_CN"),
    ],
)
def test_direct_context_rejects_noncanonical_effective_options_before_typed_load(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    options: dict[str, Any],
    style_locale: str,
) -> None:
    import docwen_plugin_markdown.to_docx.resolved_v4_route as route_module

    monkeypatch.setattr(route_module, "load_resolved_v4_inputs", _forbidden_legacy)
    context, workspace = _context(
        tmp_path,
        _refs(_NEUTRAL, _PLAN),
        options=options,
        style_locale=style_locale,
    )

    result = MdToDocxConverter().convert(context)

    assert not result.success
    assert result.error is not None
    assert result.error.diagnostic_code == "MD2DOCX-RESOLVED-V4-OPTIONS-INVALID"
    assert workspace.registered_artifacts == []
    assert list(Path(workspace.staging_dir).iterdir()) == []


@pytest.mark.parametrize("failure_stage", ["renderer", "proof"])
def test_failure_after_resource_staging_leaves_zero_staging(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    failure_stage: str,
) -> None:
    import docwen_plugin_markdown.to_docx.resolved_v4_route as route_module

    if failure_stage == "renderer":
        monkeypatch.setattr(route_module.MdToDocxRenderer, "render", _forbidden_legacy)
    else:

        def reject_proof(_self: Any, _path: Any) -> None:
            raise ResolvedNumberingDocxError("forced proof failure")

        monkeypatch.setattr(ResolvedNumberingDocxSession, "prove_package", reject_proof)
    context, workspace = _context(tmp_path, _refs(_NEUTRAL, _PLAN))

    result = MdToDocxConverter().convert(context)

    assert not result.success
    assert result.artifacts == []
    assert workspace.registered_artifacts == []
    assert list(Path(workspace.staging_dir).iterdir()) == []


def test_partial_binder_write_is_request_owned_and_fully_removed(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_plugin_markdown.to_docx.resolved_v4_route as route_module

    def partial_write_then_reject(_document: Any, resource_root: Path) -> Any:
        resource_root.mkdir()
        (resource_root / "first.png").write_bytes(b"complete-first-resource")
        (resource_root / "second.png").write_bytes(b"partial-second")
        raise route_module.ResolvedResourceStagingError("forced second-resource write failure")

    monkeypatch.setattr(route_module, "bind_resolved_document_resources", partial_write_then_reject)
    context, workspace = _context(tmp_path, _refs(_NEUTRAL, _PLAN))

    result = MdToDocxConverter().convert(context)

    assert not result.success
    assert result.error is not None
    assert result.error.diagnostic_code == "docwen.resolved_document.invalid"
    assert workspace.registered_artifacts == []
    assert list(Path(workspace.staging_dir).iterdir()) == []


def test_preexisting_resource_directory_is_rejected_without_deleting_it(tmp_path: Path) -> None:
    context, workspace = _context(tmp_path, _refs(_NEUTRAL, _PLAN))
    preexisting = Path(workspace.staging_dir) / "resolved-v4-resources"
    preexisting.mkdir()
    sentinel = preexisting / "sentinel.txt"
    sentinel.write_text("owned-before-request", encoding="utf-8")

    result = MdToDocxConverter().convert(context)

    assert not result.success
    assert result.error is not None
    assert result.error.diagnostic_code == "docwen.resolved_document.invalid"
    assert workspace.registered_artifacts == []
    assert sentinel.read_text(encoding="utf-8") == "owned-before-request"
    assert sorted(item.name for item in Path(workspace.staging_dir).iterdir()) == ["resolved-v4-resources"]


def test_preexisting_sidecar_is_rejected_without_deleting_it(tmp_path: Path) -> None:
    context, workspace = _context(tmp_path, _refs(_NEUTRAL, _PLAN))
    preexisting = Path(workspace.staging_dir) / "artifact_1.docx.docwen"
    preexisting.write_bytes(b"owned-before-request")

    result = MdToDocxConverter().convert(context)

    assert not result.success
    assert result.error is not None
    assert workspace.registered_artifacts == []
    assert preexisting.read_bytes() == b"owned-before-request"
    assert sorted(item.name for item in Path(workspace.staging_dir).iterdir()) == [preexisting.name]


@pytest.mark.parametrize("failure_stage", ["renderer", "success_cleanup"])
def test_cleanup_error_never_masks_failure_or_returns_success(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    failure_stage: str,
) -> None:
    import docwen_plugin_markdown.to_docx.resolved_v4_route as route_module

    original_cleanup = route_module._remove_request_resource_root

    def cleanup_then_reject(staging_dir: Path, resource_root: Path) -> None:
        original_cleanup(staging_dir, resource_root)
        raise RuntimeError("forced cleanup reporting failure")

    monkeypatch.setattr(route_module, "_remove_request_resource_root", cleanup_then_reject)
    if failure_stage == "renderer":
        monkeypatch.setattr(route_module.MdToDocxRenderer, "render", _forbidden_legacy)
    context, workspace = _context(tmp_path, _refs(_NEUTRAL, _PLAN))

    result = MdToDocxConverter().convert(context)

    assert not result.success
    assert result.error is not None
    if failure_stage == "renderer":
        assert "historical source converter" in result.error.message
    else:
        assert "forced cleanup reporting failure" in result.error.message
    assert workspace.registered_artifacts == []
    assert list(Path(workspace.staging_dir).iterdir()) == []


def test_predelete_cleanup_denial_is_reported_with_the_primary_failure(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_plugin_markdown.to_docx.resolved_v4_route as route_module

    def deny_cleanup(_staging_dir: Path, _resource_root: Path) -> None:
        raise PermissionError("forced pre-delete denial")

    monkeypatch.setattr(route_module, "_remove_request_resource_root", deny_cleanup)
    monkeypatch.setattr(route_module.MdToDocxRenderer, "render", _forbidden_legacy)
    context, workspace = _context(tmp_path, _refs(_NEUTRAL, _PLAN))

    result = MdToDocxConverter().convert(context)

    assert not result.success
    assert result.error is not None
    assert "historical source converter" in result.error.message
    assert "resource cleanup failed (PermissionError)" in result.error.message
    assert [item.code for item in result.diagnostics][-1] == "MD2DOCX-RESOLVED-V4-CLEANUP-FAILED"
    assert workspace.registered_artifacts == []
    residue = Path(workspace.staging_dir) / "resolved-v4-resources"
    assert residue.is_dir()
    assert (residue / "image-system.png").is_file()
    shutil.rmtree(residue)


def test_cleanup_type_swap_is_reported_without_deleting_the_replacement(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_plugin_markdown.to_docx.resolved_v4_route as route_module

    original_cleanup = route_module._remove_request_resource_root
    swapped = False

    def swap_to_file_then_cleanup(staging_dir: Path, resource_root: Path) -> None:
        nonlocal swapped
        if not swapped:
            shutil.rmtree(resource_root)
            resource_root.write_text("foreign replacement", encoding="utf-8")
            swapped = True
        original_cleanup(staging_dir, resource_root)

    monkeypatch.setattr(route_module, "_remove_request_resource_root", swap_to_file_then_cleanup)
    monkeypatch.setattr(route_module.MdToDocxRenderer, "render", _forbidden_legacy)
    context, workspace = _context(tmp_path, _refs(_NEUTRAL, _PLAN))

    result = MdToDocxConverter().convert(context)

    assert not result.success
    assert result.error is not None
    assert "resource cleanup failed (RuntimeError)" in result.error.message
    replacement = Path(workspace.staging_dir) / "resolved-v4-resources"
    assert replacement.read_text(encoding="utf-8") == "foreign replacement"
    assert [item.code for item in result.diagnostics][-1] == "MD2DOCX-RESOLVED-V4-CLEANUP-FAILED"
    assert workspace.registered_artifacts == []
    replacement.unlink()


def test_artifact_registration_failure_removes_output_and_leaves_no_registration(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    context, workspace = _context(tmp_path, _refs(_NEUTRAL, _PLAN))

    def reject_registration(_manifest: Any) -> None:
        raise RuntimeError("forced registration failure")

    monkeypatch.setattr(workspace, "add_artifact", reject_registration)

    result = MdToDocxConverter().convert(context)

    assert not result.success
    assert result.error is not None
    assert "forced registration failure" in result.error.message
    assert workspace.registered_artifacts == []
    assert list(Path(workspace.staging_dir).iterdir()) == []


def test_rich_caption_uses_rendered_binding_and_heading_merge_snapshot(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    neutral, plan = _write_rich_pair(tmp_path)
    calls: list[str] = []
    original_rich = ResolvedNumberingDocxSession.bind_rendered_caption
    original_plain = ResolvedNumberingDocxSession.bind_caption

    def rich(self: ResolvedNumberingDocxSession, *args: Any, **kwargs: Any) -> None:
        calls.append("rich")
        original_rich(self, *args, **kwargs)

    def plain(self: ResolvedNumberingDocxSession, *args: Any, **kwargs: Any) -> None:
        calls.append("plain")
        original_plain(self, *args, **kwargs)

    monkeypatch.setattr(ResolvedNumberingDocxSession, "bind_rendered_caption", rich)
    monkeypatch.setattr(ResolvedNumberingDocxSession, "bind_caption", plain)
    context, _workspace = _context(
        tmp_path / "run",
        _refs(neutral, plan),
        options={"locale": "zh_CN", "heading_merge_mode": "always"},
    )

    result = MdToDocxConverter().convert(context)

    assert result.success, result.error
    assert calls == ["rich"]
    with ZipFile(result.artifacts[0].staging_path) as package:
        document = etree.fromstring(package.read("word/document.xml"))
    instructions = [item.text or "" for item in document.iter(qn("w:instrText"))]
    visible = "".join(item.text or "" for item in document.iter(qn("w:t")))
    assert sum(" REF " in item for item in instructions) == 1
    assert sum(" CITATION " in item for item in instructions) == 1
    assert "2.3 Authored:Merged body." in visible
    assert visible.count("@[[#^head-a|Heading]]") == 1


def _write_rich_pair(tmp_path: Path) -> tuple[Path, Path]:
    token = "@[[#^head-a|Heading]]"
    citation_token = "@cite"
    caption_text = f"literal {token} then {token} and {citation_token}"
    source = f"# 2.3 Authored: ^head-a\n\nMerged body.\n\nTable: {caption_text} ^table-a\n| A |\n|---|\n| B |\n"
    heading_end = source.index("\n")
    table_start = source.index("Table:")
    table_end = len(source) - 1
    reference_start = source.index(token, source.index(token) + 1)
    citation_start = source.index(citation_token)
    targets = (
        ResolvedDocumentTarget(
            0,
            heading_end,
            _sha_text(source[:heading_end]),
            "heading",
            "head-a",
            1,
            "2.3 Authored:",
        ),
        ResolvedDocumentTarget(
            table_start,
            table_end,
            _sha_text(source[table_start:table_end]),
            "table",
            "table-a",
            None,
            caption_text,
        ),
    )
    reference = ResolvedReference(
        reference_start,
        reference_start + len(token),
        _sha_text(token),
        token,
        0,
        heading_end,
        "heading",
        "head-a",
        "1",
        "Heading",
    )
    citation = ResolvedCitation(
        citation_start,
        citation_start + len(citation_token),
        _sha_text(citation_token),
        citation_token,
        "narrative",
        "cluster-a",
        (ResolvedCitationItem("cite", "reference-record:98", _sha_text("record"), "Citation"),),
        "Citation",
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
        (definition,),
        (HeadingInstance("document", "main", ()),),
        (
            NumberingTarget(
                0,
                heading_end,
                "heading",
                True,
                "head-a",
                "1",
                HeadingListMaterialization("main", "document", 1),
            ),
            NumberingTarget(
                table_start,
                table_end,
                "table",
                True,
                "table-a",
                "1",
                table_materialization,
            ),
        ),
    )
    plan_body = json.loads(json.dumps(asdict(plan)))
    plan_body["heading_definitions"][0]["levels"][0]["display"] = [
        {"counter": {"level": 1, "number_format": "arabic_half"}}
    ]
    plan_sha256 = hashlib.sha256(canonicalize_numbering_plan(plan_body)).hexdigest()
    source_sha256 = _sha_text(source)
    document = ResolvedDocument(source, targets, (reference,), (), (citation,), ())
    neutral_payload = {
        "$schema": "urn:docwen:schema:resolved-document:v1",
        "schema": "docwen.resolved_document.v1",
        "input_id": "rich-document",
        "source_sha256": source_sha256,
        "plan_sha256": plan_sha256,
        "document": asdict(document),
    }
    plan_payload = {
        "$schema": "urn:docwen:schema:numbering-export-plan:v1",
        "schema": "docwen.numbering_export_plan.v1",
        "input_id": "rich-document",
        "source_sha256": source_sha256,
        "plan_sha256": plan_sha256,
        "plan": plan_body,
    }
    neutral_path = tmp_path / "rich-neutral.json"
    plan_path = tmp_path / "rich-plan.json"
    neutral_path.write_text(json.dumps(neutral_payload, separators=(",", ":")), encoding="utf-8")
    plan_path.write_text(json.dumps(plan_payload, separators=(",", ":")), encoding="utf-8")
    return neutral_path, plan_path
