"""Machine markup discovery preserves real runtime dependency gates."""

from copy import deepcopy
from pathlib import Path
from typing import Any

import pytest

from docwen_application.conversion_contracts import ConversionServiceError
from docwen_application.conversion_service import ConversionService
from docwen_core.formats import get_media_type
from docwen_plugin_markup.manifest import build_manifest
from docwen_runtime import capabilities

from ._conversion_service_support import _Committer, _Controller, _request

pytestmark = pytest.mark.integration


class _MarkupController(_Controller):
    def __init__(self, projection: dict[str, Any]) -> None:
        super().__init__()
        self.projection = projection

    def describe_runtime_capabilities(self) -> dict[str, Any]:
        return deepcopy(self.projection)


@pytest.mark.parametrize("source", ["html", "enex"])
@pytest.mark.parametrize("markdownify_available", [False, True])
def test_markup_machine_capability_tracks_required_and_optional_dependencies(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, source: str, markdownify_available: bool
) -> None:
    monkeypatch.setattr(
        capabilities,
        "_gate_status",
        lambda gate: {"id": gate, "available": markdownify_available if gate == "python.markdownify" else False},
    )
    projection = capabilities.build_runtime_capability_projection([build_manifest()], platform_id="windows")
    service = ConversionService(_MarkupController(projection), _Committer())
    [capability] = [
        item for item in service.list_capabilities() if item.capability_id == f"convert.{source}.to_markdown"
    ]
    assert (capability.availability != "unavailable") is markdownify_available
    assert {"dependency_id": "python.markdownify", "required": True, "available": markdownify_available} in (
        capability.dependencies
    )
    assert {"dependency_id": "python.rapidocr", "required": False, "available": False} in capability.dependencies
    assert capability.input_shape.slots[0].media_types == (get_media_type(source),)
    assert capability.output_media_types == ("text/markdown",)
    assert capability.output_shape.artifact_kinds == ("document", "fragment", "resource")
    assert capability.output_shape.relation_types == ("fragment_of", "resource_of")
    assert set(capability.options_schema["properties"]) == {
        "recognize_text",
        "preserve_resources",
        "recognize_tables",
        "ocr_language",
        "image_mode",
        "image_link_style",
        "ocr_placement",
    }
    source_path = tmp_path / ("source." + source)
    source_path.write_text("synthetic handle", encoding="utf-8")
    staging = tmp_path / "staging"
    staging.mkdir()
    request = _request(
        source_path,
        staging,
        capability_id=capability.capability_id,
        media_type=get_media_type(source),
        options={"recognize_text": False, "preserve_resources": True},
    )
    if markdownify_available:
        plan = service.plan(request)
        assert plan.output_shape == capability.output_shape
    else:
        with pytest.raises(ConversionServiceError) as caught:
            service.plan(request)
        assert caught.value.code == "capability_unavailable"
        assert caught.value.details["missing_required_dependencies"] == ["python.markdownify"]
