"""Real Machine markup conversion retains document and embedded-resource semantics."""

from __future__ import annotations

import base64
import hashlib
from io import BytesIO
from pathlib import Path
from typing import Any

import pytest
from docx import Document

from docwen_application.controller import ApplicationController
from docwen_application.conversion_contracts import ConversionPlanRequest, LocalInputHandle, StagingOutputTarget
from docwen_application.conversion_service import ConversionService
from docwen_core.formats import get_media_type
from docwen_core.paths import filesystem_path
from docwen_runtime.output.artifact_bundle import ArtifactBundleCommitter

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]


@pytest.mark.parametrize("source_format", ["html", "enex"])
def test_machine_markup_returns_markdown_and_original_embedded_bytes(
    round_trip_runtime: Any, tmp_path: Path, source_format: str
) -> None:
    png = base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
    )
    expected_resources = [png]
    png64 = base64.b64encode(png).decode("ascii")
    if source_format == "html":
        content = f'<!doctype html><html><body><h1>Markup title</h1><p>Body marker</p><img src="data:image/png;base64,{png64}"></body></html>'
    else:
        buffer = BytesIO()
        doc = Document()
        doc.add_paragraph("Embedded document marker")
        doc.save(buffer)
        attachment = buffer.getvalue()
        expected_resources.append(attachment)
        resources = ""
        media = ""
        for name, mime, payload in [
            ("marker.png", "image/png", png),
            ("attachment.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", attachment),
        ]:
            digest = hashlib.md5(payload, usedforsecurity=False).hexdigest()
            media += f'<en-media type="{mime}" hash="{digest}"/>'
            resources += (
                f'<resource><data encoding="base64">{base64.b64encode(payload).decode("ascii")}</data>'
                f"<mime>{mime}</mime><resource-attributes><file-name>{name}</file-name></resource-attributes></resource>"
            )
        content = (
            '<?xml version="1.0" encoding="UTF-8"?><en-export><note><title>Markup title</title>'
            f"<content><![CDATA[<en-note><div>Body marker</div>{media}</en-note>]]></content>{resources}</note></en-export>"
        )
    source = tmp_path / ("authored." + source_format)
    original = content.encode("utf-8")
    source.write_bytes(original)
    staging = tmp_path / "output"
    staging.mkdir()
    controller = ApplicationController(runtime_port=round_trip_runtime)
    controller.start()
    service = ConversionService(controller, ArtifactBundleCommitter())
    capability = next(
        item for item in service.list_capabilities() if item.capability_id == f"convert.{source_format}.to_markdown"
    )
    assert capability.availability != "unavailable", capability.dependencies
    request = ConversionPlanRequest(
        capability_id=capability.capability_id,
        inputs=(
            LocalInputHandle(
                "source.1",
                str(source),
                get_media_type(source_format),
                len(original),
                hashlib.sha256(original).hexdigest(),
                capability.input_shape.slots[0].kind,
                "source",
                source.name,
            ),
        ),
        output=StagingOutputTarget(str(staging)),
        options={"recognize_text": False, "preserve_resources": True, "image_link_style": "markdown_embed"},
    )
    plan = service.plan(request)
    outcome = service.execute_accepted(service.accept(plan.plan_id))
    assert outcome.state == "completed", outcome.error
    assert outcome.bundle is not None
    bundle = outcome.bundle
    [document] = [item for item in bundle.artifacts if item.kind == "document"]
    assert document.media_type == "text/markdown"
    markdown = filesystem_path(staging / document.locator).read_text(encoding="utf-8")
    assert "Markup title" in markdown and "Body marker" in markdown
    resources = [item for item in bundle.artifacts if item.kind == "resource"]
    assert len(resources) == len(expected_resources)
    assert {filesystem_path(staging / item.locator).read_bytes() for item in resources} == set(expected_resources)
    assert all(
        item.type == "resource_of" and item.target_artifact_id == document.artifact_id for item in bundle.relations
    )
    assert {item.source_artifact_id for item in bundle.relations} == {item.artifact_id for item in resources}
    for resource in resources:
        assert Path(resource.locator).name in markdown
    assert source.read_bytes() == original
