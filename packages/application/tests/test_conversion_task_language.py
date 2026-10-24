"""Machine language ownership through planning, admission and real DOCX output."""

from __future__ import annotations

from pathlib import Path
from typing import Any
from unittest.mock import Mock

import pytest
from docx import Document

from docwen_application.controller import ApplicationController
from docwen_application.conversion_contracts import (
    MARKDOWN_SOURCE_TO_DOCX_CAPABILITY_ID,
    MARKDOWN_TO_DOCX_CAPABILITY_ID,
)
from docwen_application.conversion_service import ConversionService
from docwen_application.ports.runtime import ConfigPort
from docwen_core.cancellation import CancellationToken
from docwen_core.export_semantics import MarkdownExportSemantics
from docwen_plugin_markdown.plugin import MarkdownPlugin
from docwen_runtime._execution_context import RuntimeExecutionContext
from docwen_runtime.adapters import RuntimePortAdapter
from docwen_runtime.config.document_styles import build_document_style_catalog
from docwen_runtime.workspace.manager import WorkspaceHandle

from ._conversion_service_support import _Committer, _Controller, _request

pytestmark = pytest.mark.integration


@pytest.mark.parametrize("capability_id", [MARKDOWN_SOURCE_TO_DOCX_CAPABILITY_ID, MARKDOWN_TO_DOCX_CAPABILITY_ID])
@pytest.mark.parametrize(
    ("core_locale", "ui_locale", "explicit_locale", "expected_locale", "expected_style"),
    [
        ("en_US", "ja_JP", None, "en_US", "Body Paragraph"),
        ("en_US", "en_US", None, "en_US", "Body Paragraph"),
        ("ja_JP", "en_US", None, "ja_JP", "本文段落"),
        ("en_US", "en_US", "ja_JP", "ja_JP", "本文段落"),
        ("ja_JP", "ja_JP", "en_US", "en_US", "Body Paragraph"),
        (None, "ja_JP", None, "zh_CN", "正文段落"),
    ],
)
def test_machine_docx_task_uses_frozen_output_language(
    tmp_path: Path,
    capability_id: str,
    core_locale: str | None,
    ui_locale: str,
    explicit_locale: str | None,
    expected_locale: str,
    expected_style: str,
) -> None:
    source = tmp_path / "note.md"
    source.write_text("Body text.\n", encoding="utf-8")
    staging = tmp_path / "plan-staging"
    staging.mkdir()
    options = {} if explicit_locale is None else {"locale": explicit_locale}
    request = _request(source, staging, capability_id=capability_id, options=options)
    source_images = {item.path: Path(item.path).read_bytes() for item in request.inputs}
    live_config: dict[str, Any] = {"gui": {"language": {"locale": ui_locale}}}
    if core_locale is not None:
        live_config["conversion"] = {"output_locale": core_locale}
    config = Mock(spec=ConfigPort)
    config.snapshot.side_effect = lambda: live_config
    rendered = []

    def render(admitted: Any, *, on_event: Any = None) -> Any:
        assert admitted.options.get("locale") == explicit_locale
        if explicit_locale is None:
            assert "locale" not in admitted.options
        expected_roles = (
            ["source"]
            if capability_id == MARKDOWN_SOURCE_TO_DOCX_CAPABILITY_ID
            else ["neutral_document", "numbering_export_plan"]
        )
        assert [ref.input_role for ref in admitted.input_refs] == expected_roles
        assert admitted.config_snapshot == live_config
        assert admitted.config_snapshot is not live_config
        # A live configuration change after controller/adapter admission cannot
        # alter the language selected by this task's frozen snapshot.
        live_config["conversion"] = {"output_locale": "de_DE"}
        live_config["gui"]["language"]["locale"] = "fr_FR"
        catalog = build_document_style_catalog(admitted.config_snapshot, request_options=admitted.options)
        assert catalog.locale == expected_locale
        workspace = WorkspaceHandle(
            admitted.input_refs[0].path, admitted.output_policy.output_dir, tuple(admitted.input_refs)
        )
        context = RuntimeExecutionContext(
            admitted,
            workspace,
            admitted.config_snapshot,
            CancellationToken(),
            on_event=on_event,
            document_style_catalog=catalog,
            markdown_export_semantics=MarkdownExportSemantics.from_config_snapshot(
                admitted.config_snapshot, requested_locale=admitted.options.get("locale")
            ),
        )
        expected_labels = {
            "en_US": {"title": "title", "subtitle": "subtitle"},
            "ja_JP": {"title": "タイトル", "subtitle": "サブタイトル"},
            "zh_CN": {"title": "标题", "subtitle": "副标题"},
        }
        assert context.yaml_key_labels == expected_labels[expected_locale]
        result = MarkdownPlugin().convert(context)
        rendered.append(result)
        return result

    # TaskManager execution, workspace setup, capability discovery and Bundle
    # commit use substitutes. Service, controller, adapter admission, language
    # builders, the context/handle classes, plugin and DOCX renderers are real.
    manager = Mock()
    manager.execute_single.side_effect = render
    manager.execute_batch.side_effect = AssertionError("DOCX typed inputs must execute as one task")
    adapter = RuntimePortAdapter(manager, capability_provider=_Controller().describe_runtime_capabilities)
    controller = ApplicationController(runtime_port=adapter, config_port=config)
    service = ConversionService(controller, _Committer())
    plan = service.plan(request)
    assert plan.effective_options.get("locale") == explicit_locale
    if explicit_locale is None:
        assert "locale" not in plan.effective_options
    config.snapshot.assert_not_called()
    # Caller-visible values are projections, not the accepted task's authority.
    request.options["locale"] = "de_DE"
    plan.effective_options["locale"] = "fr_FR"
    task_id = service.accept(plan.plan_id)
    config.snapshot.assert_not_called()
    outcome = service.execute_accepted(task_id)
    assert outcome.state == "completed", outcome
    config.snapshot.assert_called_once_with()
    manager.execute_single.assert_called_once()
    manager.execute_batch.assert_not_called()
    result = rendered[0]
    assert result.success, result.error
    document = Document(result.artifacts[0].staging_path)
    style = next(style for style in document.styles if style.style_id == "DocWenBodyParagraph")
    assert style.name == expected_style
    paragraph = next(paragraph for paragraph in document.paragraphs if paragraph.text == "Body text.")
    assert paragraph.style is not None
    assert paragraph.style.style_id == "DocWenBodyParagraph"
    assert all(Path(path).read_bytes() == value for path, value in source_images.items())
