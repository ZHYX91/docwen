"""Closed exact-two input and projection gates for the Markdown plugin."""

from __future__ import annotations

import base64
import hashlib
import json
from pathlib import Path

import pytest
from tests.support.workspace import FakeWorkspaceHandle

from docwen_core.models.file_ref import FileRef
from docwen_core.models.resolved_numbering import (
    NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
    RESOLVED_DOCUMENT_MEDIA_TYPE,
    NumberingExportPlanEnvelope,
    NumberingTarget,
    ResolvedDocument,
    ResolvedDocumentEnvelope,
    ResolvedDocumentTarget,
    ResolvedEmbeddedResource,
    ResolvedNumberingPlan,
    ResolvedNumberingPort,
    ResolvedReference,
    ResolvedResourceOccurrence,
    canonicalize_numbering_plan,
)
from docwen_core.resolved_resource_staging import (
    ResolvedResourceBinding,
    ResolvedTextEdit,
    ResolvedTextProjection,
    bind_resolved_document_resources,
)
from docwen_plugin_markdown.mistune_extensions import parse_markdown_text
from docwen_plugin_markdown.resolved_conversion_v4 import (
    PreparedResolvedInputsV4,
    ResolvedConversionV4Unsupported,
    claims_resolved_v4_inputs,
    compose_resolved_v4_markdown,
    load_resolved_v4_inputs,
    prove_resolved_v4_image_inventory,
)
from docwen_plugin_markdown.resolved_runtime_v4 import (
    apply_resolved_runtime_v4,
    prepare_resolved_runtime_v4,
)
from docwen_plugin_markdown.resolved_source_carriers_v4 import (
    apply_resolved_source_carriers_v4,
    prepare_resolved_source_carriers_v4,
)

pytestmark = pytest.mark.unit

_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEklEQVR4nGNUSFjAwMDAxAAGAA0qASTlOPBgAAAAAElFTkSuQmCC"
)


def _sha_text(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def _sha_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def _write_empty_pair(tmp_path: Path) -> tuple[Path, Path]:
    plan_body = {"heading_definitions": [], "heading_instances": [], "targets": []}
    plan_sha256 = _sha_bytes(canonicalize_numbering_plan(plan_body))
    source_sha256 = _sha_text("")
    neutral_payload = {
        "$schema": "urn:docwen:schema:resolved-document:v1",
        "schema": "docwen.resolved_document.v1",
        "input_id": "document-1",
        "source_sha256": source_sha256,
        "plan_sha256": plan_sha256,
        "document": {
            "authored_markdown": "",
            "targets": [],
            "references": [],
            "resource_occurrences": [],
            "citations": [],
            "resources": [],
        },
    }
    plan_payload = {
        "$schema": "urn:docwen:schema:numbering-export-plan:v1",
        "schema": "docwen.numbering_export_plan.v1",
        "input_id": "document-1",
        "source_sha256": source_sha256,
        "plan_sha256": plan_sha256,
        "plan": plan_body,
    }
    neutral = tmp_path / "neutral.json"
    plan = tmp_path / "plan.json"
    neutral.write_text(json.dumps(neutral_payload, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    plan.write_text(json.dumps(plan_payload, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    return neutral, plan


def _ref(path: Path, *, role: str) -> FileRef:
    is_document = role == "neutral_document"
    return FileRef(
        path=str(path),
        format="markdown" if is_document else "json",
        category="document" if is_document else "resource",
        size_bytes=path.stat().st_size,
        input_kind="document" if is_document else "resource",
        input_role=role,
        logical_path="document.json" if is_document else "numbering-plan.json",
        media_type=RESOLVED_DOCUMENT_MEDIA_TYPE if is_document else NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
    )


def _workspace(tmp_path: Path, refs: tuple[FileRef, ...]) -> FakeWorkspaceHandle:
    staging = tmp_path / "staging"
    staging.mkdir(exist_ok=True)
    return FakeWorkspaceHandle(refs[0].path if refs else "", str(staging), refs)


def _direct_port(source: str, *, include_unbound_image: bool = False) -> ResolvedNumberingPort:
    heading_line = "# Heading ^heading-one"
    heading_end = len(heading_line)
    image_token = "![pixel](pixel.png)"
    image_start = source.index(image_token)
    reference_token = "@[[#^heading-one]]"
    reference_start = source.index(reference_token)
    target = ResolvedDocumentTarget(
        source_start=0,
        source_end=heading_end,
        source_slice_sha256=_sha_text(source[:heading_end]),
        kind="heading",
        target_id="heading-one",
        heading_level=1,
        authored_text="Heading",
    )
    reference = ResolvedReference(
        source_start=reference_start,
        source_end=reference_start + len(reference_token),
        source_slice_sha256=_sha_text(reference_token),
        authored_token=reference_token,
        target_source_start=0,
        target_source_end=heading_end,
        target_kind="heading",
        target_id="heading-one",
        cached_number="1",
        alias=None,
    )
    resource = ResolvedEmbeddedResource(
        resource_id="image-one",
        role="linked_resource",
        media_type="image/png",
        size_bytes=len(_PNG),
        sha256=_sha_bytes(_PNG),
        content=_PNG,
    )
    occurrence = ResolvedResourceOccurrence(
        source_start=image_start,
        source_end=image_start + len(image_token),
        source_slice_sha256=_sha_text(image_token),
        authored_token=image_token,
        authored_locator="pixel.png",
        resource_id="image-one",
    )
    source_sha256 = _sha_text(source)
    plan_target = NumberingTarget(
        source_start=0,
        source_end=heading_end,
        kind="heading",
        enabled=False,
        target_id="heading-one",
        derived_number=None,
        materialization=None,
    )
    plan = ResolvedNumberingPlan(heading_definitions=(), heading_instances=(), targets=(plan_target,))
    plan_sha256 = "1" * 64
    document = ResolvedDocument(
        authored_markdown=source,
        targets=(target,),
        references=(reference,),
        resource_occurrences=(occurrence,),
        citations=(),
        resources=(resource,),
    )
    assert include_unbound_image or source.count("![") == 1
    return ResolvedNumberingPort(
        ResolvedDocumentEnvelope("document-1", source_sha256, plan_sha256, document),
        NumberingExportPlanEnvelope("document-1", source_sha256, plan_sha256, plan),
    )


def _prepared(port: ResolvedNumberingPort) -> PreparedResolvedInputsV4:
    return PreparedResolvedInputsV4(
        port=port,
        runtime_plan=prepare_resolved_runtime_v4(port),
        source_carrier_plan=prepare_resolved_source_carriers_v4(
            port.document.authored_markdown,
            input_id=port.input_id,
            expected_source_sha256=port.source_sha256,
        ),
        neutral_document_path=Path("neutral.json"),
        numbering_export_plan_path=Path("plan.json"),
    )


def _walk_ast(nodes: list[dict[str, object]]):
    for node in nodes:
        yield node
        children = node.get("children")
        if isinstance(children, list):
            yield from _walk_ast(children)


def test_exact_pair_is_reloaded_from_workspace_copies(tmp_path: Path) -> None:
    neutral, plan = _write_empty_pair(tmp_path)
    refs = (_ref(neutral, role="neutral_document"), _ref(plan, role="numbering_export_plan"))

    prepared = load_resolved_v4_inputs(_workspace(tmp_path, refs))

    assert prepared.port.document.authored_markdown == ""
    assert prepared.neutral_document_path == neutral
    assert prepared.numbering_export_plan_path == plan


def test_partial_or_mixed_resolved_claim_never_falls_through(tmp_path: Path) -> None:
    neutral, plan = _write_empty_pair(tmp_path)
    neutral_ref = _ref(neutral, role="neutral_document")
    plan_ref = _ref(plan, role="numbering_export_plan")
    legacy = FileRef(path=str(neutral), format="markdown", category="document", input_role="source")
    assert claims_resolved_v4_inputs((legacy,)) is False
    assert claims_resolved_v4_inputs((neutral_ref,)) is True

    with pytest.raises(ResolvedConversionV4Unsupported, match="one exact input pair") as missing:
        load_resolved_v4_inputs(_workspace(tmp_path, (neutral_ref,)))
    assert missing.value.code == "docwen.numbering_export_plan.missing"

    with pytest.raises(ResolvedConversionV4Unsupported, match="one exact input pair") as extra:
        load_resolved_v4_inputs(_workspace(tmp_path, (neutral_ref, plan_ref, legacy)))
    assert extra.value.code == "docwen.resolved_document.invalid"


@pytest.mark.parametrize(
    ("role", "attribute", "value", "code"),
    [
        ("neutral_document", "input_kind", "resource", "docwen.resolved_document.invalid"),
        ("neutral_document", "media_type", "application/json", "docwen.resolved_document.invalid"),
        ("numbering_export_plan", "input_kind", "document", "docwen.numbering_export_plan.invalid"),
        ("numbering_export_plan", "media_type", "application/json", "docwen.numbering_export_plan.invalid"),
    ],
)
def test_exact_pair_metadata_is_rechecked(
    tmp_path: Path,
    role: str,
    attribute: str,
    value: str,
    code: str,
) -> None:
    neutral, plan = _write_empty_pair(tmp_path)
    refs = [_ref(neutral, role="neutral_document"), _ref(plan, role="numbering_export_plan")]
    item = next(candidate for candidate in refs if candidate.input_role == role)
    setattr(item, attribute, value)

    with pytest.raises(ResolvedConversionV4Unsupported) as rejected:
        load_resolved_v4_inputs(_workspace(tmp_path, tuple(refs)))
    assert rejected.value.code == code


def test_markers_and_private_resources_compose_by_original_ranges(tmp_path: Path) -> None:
    source = "# Heading ^heading-one\n\n![pixel](pixel.png)\n\nSee @[[#^heading-one]].\n"
    port = _direct_port(source)
    binding = bind_resolved_document_resources(port.document, tmp_path / "resources")

    projection = compose_resolved_v4_markdown(_prepared(port), binding)
    ast = parse_markdown_text(projection.markdown)
    restored = apply_resolved_runtime_v4(ast, projection.runtime_plan)
    prove_resolved_v4_image_inventory(restored, projection)

    assert len(projection.expected_image_urls) == 1
    assert projection.expected_image_urls[0].endswith("/resources/image-one.png")
    assert "pixel.png" not in projection.markdown
    assert "@[[#^heading-one]]" not in projection.markdown


def test_resolved_resource_and_source_carrier_edits_share_original_coordinates(tmp_path: Path) -> None:
    source = (
        "# Heading ^heading-one\n\n"
        "![pixel](pixel.png) ^raw-image\n\n"
        "See @[[#^heading-one]].\n\n"
        "> ```mermaid\n"
        "> graph TD\n"
        "> ```\n"
        ">\n"
        "> ^inner-fence\n\n"
        "^outer-quote\n"
    )
    port = _direct_port(source)
    binding = bind_resolved_document_resources(port.document, tmp_path / "resources")

    projection = compose_resolved_v4_markdown(_prepared(port), binding)
    ast = parse_markdown_text(projection.markdown, auto_link_bare_url=False)
    ast = apply_resolved_runtime_v4(ast, projection.runtime_plan)
    restored = apply_resolved_source_carriers_v4(ast, projection.source_carrier_plan)
    nodes = tuple(_walk_ast(restored))

    image_paragraph = next(
        item for item in nodes if item.get("type") == "paragraph" and "_docwen_v3_ordinary_anchor" in item
    )
    assert image_paragraph["_docwen_v3_ordinary_anchor"]["id"] == "raw-image"  # type: ignore[index]
    fence = next(item for item in nodes if item.get("type") == "block_code")
    assert fence["_docwen_v3_fenced_body"] == "graph TD\n"
    assert fence["_docwen_v3_ordinary_anchor"]["id"] == "inner-fence"  # type: ignore[index]
    quote = next(item for item in nodes if item.get("type") == "block_quote")
    assert quote["_docwen_v3_ordinary_anchor"]["id"] == "outer-quote"  # type: ignore[index]
    prove_resolved_v4_image_inventory(restored, projection)  # type: ignore[arg-type]


def test_semantic_and_resource_edit_overlap_is_rejected(tmp_path: Path) -> None:
    source = "# Heading ^heading-one\n\n![pixel](pixel.png)\n\nSee @[[#^heading-one]].\n"
    port = _direct_port(source)
    prepared = _prepared(port)
    binding = bind_resolved_document_resources(port.document, tmp_path / "resources")
    reference = port.document.references[0]
    replacement = "![forged](<D:/request/forged.png>)"
    edit = ResolvedTextEdit(
        source_start=reference.source_start,
        source_end=reference.source_end,
        replacement=replacement,
        result_start=reference.source_start,
        result_end=reference.source_start + len(replacement),
    )
    rendered = source[: reference.source_start] + replacement + source[reference.source_end :]
    forged = ResolvedResourceBinding(
        rendered_markdown=rendered,
        linked_paths=binding.linked_paths,
        bibliography=binding.bibliography,
        text_projection=ResolvedTextProjection(
            source_length=len(source),
            result_length=len(rendered),
            edits=(edit,),
        ),
    )

    with pytest.raises(ResolvedConversionV4Unsupported, match="overlap"):
        compose_resolved_v4_markdown(prepared, forged)


def test_unbound_image_node_is_rejected_before_renderer_can_read_it(tmp_path: Path) -> None:
    source = (
        "# Heading ^heading-one\n\n![pixel](pixel.png)\n\n![unbound](physical-decoy.png)\n\nSee @[[#^heading-one]].\n"
    )
    port = _direct_port(source, include_unbound_image=True)
    binding = bind_resolved_document_resources(port.document, tmp_path / "resources")
    projection = compose_resolved_v4_markdown(_prepared(port), binding)
    restored = apply_resolved_runtime_v4(parse_markdown_text(projection.markdown), projection.runtime_plan)

    with pytest.raises(ResolvedConversionV4Unsupported, match="image inventory"):
        prove_resolved_v4_image_inventory(restored, projection)
