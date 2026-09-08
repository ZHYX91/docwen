"""Plain text keeps its detected identity through Markdown workflow dispatch."""

from pathlib import Path

import pytest
from docx import Document

from docwen_core.detection import inspect_file
from docwen_core.models.file_inspection import FILE_INSPECTION_METADATA_KEY
from docwen_plugin_markdown import MarkdownPlugin

from .conftest import make_context

pytestmark = pytest.mark.contract


@pytest.mark.parametrize("suffix", [".txt", ".md"])
def test_admitted_plain_text_generates_docx_without_relabeling_input(tmp_path, suffix):
    source = tmp_path / f"plain{suffix}"
    text = "Ordinary text without Markdown syntax."
    source.write_text(text, encoding="utf-8")
    inspection = inspect_file(str(source))
    assert inspection.detected_format == "txt"
    assert inspection.workflow_category == "markdown"
    context, _workspace = make_context(str(source))
    input_ref = context.request.input_refs[0]
    input_ref.format = inspection.detected_format
    input_ref.category = inspection.workflow_category
    input_ref.metadata[FILE_INSPECTION_METADATA_KEY] = inspection.to_dict()

    result = MarkdownPlugin().convert(context)

    assert result.success, result.error
    document = next(artifact for artifact in result.artifacts if Path(artifact.staging_path).suffix == ".docx")
    assert text in "\n".join(paragraph.text for paragraph in Document(document.staging_path).paragraphs)
    assert input_ref.format == "txt"
    assert source.read_text(encoding="utf-8") == text
