"""Arabic request locale survives planning, template resolution and DOCX output."""

from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Any

import pytest
from docx import Document

from docwen_application.controller import ApplicationController
from docwen_application.conversion_contracts import ConversionPlanRequest, LocalInputHandle, StagingOutputTarget
from docwen_application.conversion_service import ConversionService
from docwen_core.paths import filesystem_path
from docwen_runtime.output.artifact_bundle import ArtifactBundleCommitter
from docwen_runtime.templates import TemplateRegistry

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate]


@pytest.mark.parametrize("use_english_template", [False, True])
def test_arabic_machine_conversion_preserves_locale_and_text(
    round_trip_runtime: Any, tmp_path: Path, use_english_template: bool
) -> None:
    source = tmp_path / "تقرير English 123.md"
    original = "# مرحبا بالعالم\n\nنص عربي English 123.\n".encode()
    source.write_bytes(original)
    staging = tmp_path / "output"
    staging.mkdir()
    controller = ApplicationController(runtime_port=round_trip_runtime)
    controller.start()
    try:
        service = ConversionService(controller, ArtifactBundleCommitter())
        options = {"locale": "ar_SA"}
        if use_english_template:
            template = next(
                item
                for item in TemplateRegistry.default().list_templates("docx")
                if item.name == "English General Template"
            )
            options["template_name"] = template.id
        plan = service.plan(
            ConversionPlanRequest(
                capability_id="convert.markdown_source.to_docx",
                inputs=(
                    LocalInputHandle(
                        "source.1",
                        str(source),
                        "text/markdown",
                        len(original),
                        hashlib.sha256(original).hexdigest(),
                        "document",
                        "source",
                        source.name,
                    ),
                ),
                output=StagingOutputTarget(str(staging)),
                options=options,
            )
        )
        assert plan.effective_options["locale"] == "ar_SA"
        outcome = service.execute_accepted(service.accept(plan.plan_id))
        assert outcome.state == "completed", outcome.error
        assert outcome.bundle is not None
        [artifact] = outcome.bundle.artifacts
        output = filesystem_path(staging / artifact.locator)
        assert "تقرير English 123" in output.name
        document = Document(str(output))
        text = "\n".join(paragraph.text for paragraph in document.paragraphs)
        assert "مرحبا بالعالم" in text
        assert "نص عربي English 123." in text
        assert source.read_bytes() == original
    finally:
        controller.stop()
