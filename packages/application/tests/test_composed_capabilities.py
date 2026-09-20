"""Machine document optimizers bind actual manifests, parameters and complete chains."""

from __future__ import annotations

from copy import deepcopy
from dataclasses import replace
from pathlib import Path
from typing import Any

import pytest

from docwen_application.conversion_contracts import ConversionServiceError
from docwen_application.conversion_service import ConversionService
from docwen_core.formats import get_media_type
from docwen_core.models import ArtifactManifest
from docwen_plugin_document.manifest import build_manifest as document_manifest
from docwen_plugin_optimizer_gongwen.manifest import build_manifest as gongwen_manifest
from docwen_runtime import capabilities

from ._conversion_service_support import _Committer, _Controller, _request

pytestmark = pytest.mark.integration


class _OptimizerController(_Controller):
    def __init__(self, projection: dict[str, Any]) -> None:
        super().__init__()
        self.projection = projection
        self.prepared: list[Any] = []

    def describe_runtime_capabilities(self) -> dict[str, Any]:
        return deepcopy(self.projection)

    def prepare_execution_cancellation(self, request: Any, *, batch: bool = False) -> object:
        self.prepared.append(request)
        return super().prepare_execution_cancellation(request, batch=batch)


def _projection(monkeypatch: pytest.MonkeyPatch, *, office_available: bool = True) -> dict[str, Any]:
    monkeypatch.setattr(
        capabilities,
        "_gate_status",
        lambda gate: {"id": gate, "available": office_available if gate == "external_office.word" else True},
    )
    return capabilities.build_runtime_capability_projection(
        [document_manifest(), gongwen_manifest()], platform_id="windows"
    )


@pytest.mark.parametrize("source", ["docx", "doc", "wps", "rtf", "odt"])
@pytest.mark.parametrize("office_available", [False, True])
def test_optimizer_and_plain_conversion_share_complete_chain_availability(
    monkeypatch: pytest.MonkeyPatch, source: str, office_available: bool
) -> None:
    service = ConversionService(
        _OptimizerController(_projection(monkeypatch, office_available=office_available)), _Committer()
    )
    catalog = {item.capability_id: item for item in service.list_capabilities()}
    plain = catalog[f"convert.{source}.to_markdown"]
    optimized = catalog[f"optimize.gongwen.{source}.to_markdown"]
    assert (plain.availability != "unavailable") is (source == "docx" or office_available)
    assert (optimized.availability != "unavailable") is (source == "docx" or office_available)
    assert optimized.optimization_id == "gongwen" and optimized.operation == "transform"
    assert optimized.input_shape.slots[0].media_types == (get_media_type(source),)
    if source != "docx":
        assert {
            "dependency_id": "external_office.word",
            "required": True,
            "available": office_available,
        } in optimized.dependencies
    assert "attachment_of" in optimized.output_shape.relation_types


@pytest.mark.parametrize("source", ["docx", "doc", "wps", "rtf", "odt"])
def test_plan_preserves_admitted_source_and_uses_declared_optimizer_options(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, source: str
) -> None:
    source_path = tmp_path / f"source.{source}"
    source_path.write_bytes(b"admission handle fixture")
    staging = tmp_path / "staging"
    staging.mkdir()
    controller = _OptimizerController(_projection(monkeypatch))
    service = ConversionService(controller, _Committer())
    request = _request(
        source_path,
        staging,
        capability_id=f"optimize.gongwen.{source}.to_markdown",
        media_type=get_media_type(source),
        options={"recognize_text": False, "preserve_resources": True, "image_link_style": "markdown_embed"},
    )
    request = replace(request, inputs=(replace(request.inputs[0], kind="document"),))
    plan = service.plan(request)
    service.accept(plan.plan_id)
    prepared = controller.prepared[-1]
    assert prepared.action_name == "gongwen"
    assert prepared.input_refs[0].format == source
    assert prepared.input_refs[0].path == str(source_path)
    assert prepared.options["to_md_enable_ocr"] is False
    assert prepared.options["to_md_keep_images"] is True
    assert "recognize_text" not in prepared.options
    assert "markdown_extensions" not in prepared.options
    assert plan.effective_options["numbering_scheme"] == "gongwen_standard"


def test_optimizer_rejects_options_from_the_plain_route_and_invalid_enums(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    source = tmp_path / "source.docx"
    source.write_bytes(b"admission handle fixture")
    staging = tmp_path / "staging"
    staging.mkdir()
    service = ConversionService(_OptimizerController(_projection(monkeypatch)), _Committer())
    for options, code in (
        ({"ocr_placement": "image_md"}, "unsupported_options"),
        ({"ocr_language": "chi_sim"}, "option_value_invalid"),
    ):
        with pytest.raises(ConversionServiceError) as error:
            service.plan(
                _request(
                    source,
                    staging,
                    capability_id="optimize.gongwen.docx.to_markdown",
                    media_type=get_media_type("docx"),
                    options=options,
                )
            )
        assert error.value.code == code


def test_office_disappearing_before_acceptance_prevents_execution(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    source = tmp_path / "source.odt"
    source.write_bytes(b"admission handle fixture")
    staging = tmp_path / "staging"
    staging.mkdir()
    controller = _OptimizerController(_projection(monkeypatch))
    service = ConversionService(controller, _Committer())
    request = _request(
        source, staging, capability_id="optimize.gongwen.odt.to_markdown", media_type=get_media_type("odt")
    )
    plan = service.plan(replace(request, inputs=(replace(request.inputs[0], kind="document"),)))
    controller.projection = _projection(monkeypatch, office_available=False)
    with pytest.raises(ConversionServiceError) as error:
        service.accept(plan.plan_id)
    assert error.value.code == "capability_unavailable"
    assert error.value.details == {"missing_required_dependencies": ["external_office.word"]}
    assert not controller.prepared


def test_optimizer_attachment_is_a_document_relation(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    class Attachments(_OptimizerController):
        def execute_single(self, request: Any) -> Any:
            result = super().execute_single(request)
            attachment = Path(request.output_policy.output_dir) / "attachment.md"
            attachment.write_text("# Attachment\n", encoding="utf-8")
            return replace(
                result,
                artifacts=[
                    *result.artifacts,
                    ArtifactManifest(
                        artifact_id="artifact.attachment",
                        kind="auxiliary",
                        staging_path=str(attachment),
                        suggested_name="attachment.md",
                        media_type="text/markdown",
                        metadata={"source_kind": "gongwen_attachment", "attachment_ordinal": 1},
                    ),
                ],
            )

    source = tmp_path / "source.docx"
    source.write_bytes(b"admission handle fixture")
    staging = tmp_path / "staging"
    staging.mkdir()
    service = ConversionService(Attachments(_projection(monkeypatch)), _Committer())
    plan = service.plan(
        _request(source, staging, capability_id="optimize.gongwen.docx.to_markdown", media_type=get_media_type("docx"))
    )
    outcome = service.execute_accepted(service.accept(plan.plan_id))
    assert outcome.state == "completed" and outcome.bundle is not None
    assert outcome.bundle.relations[0].type == "attachment_of"
    assert {artifact.kind for artifact in outcome.bundle.artifacts} == {"document"}
