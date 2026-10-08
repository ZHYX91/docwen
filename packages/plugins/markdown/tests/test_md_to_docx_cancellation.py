"""Cancellation must stop preprocessing before the next expensive operation."""

from pathlib import Path
from unittest.mock import Mock

import pytest

from docwen_core.cancellation import CancellationToken
from docwen_core.errors import CancellationRequested
from docwen_plugin_markdown.to_docx import converter

from .conftest import make_context, write_temp_md

pytestmark = pytest.mark.contract


@pytest.mark.parametrize(
    ("completed", "following"),
    [
        ("read_input_markdown", "prepare_runtime_semantics_v3"),
        ("prepare_runtime_semantics_v3", "resolve_template"),
        ("_markdown_image_alt_texts", "process_markdown_links"),
        ("process_markdown_links", "materialize_image_placeholders"),
        ("normalize_note_syntax", "parse_markdown_text"),
        ("apply_runtime_semantics_v3", "annotate_ast_with_merges"),
        ("extract_notes_from_ast", "analyze_document_semantics"),
        ("analyze_document_semantics", "complete_managed_styles"),
    ],
)
def test_md_to_docx_cancel_between_preprocessing_operations(monkeypatch, completed, following):
    source = Path(write_temp_md("# Heading\n\nPlain paragraph.\n"))
    original = source.read_bytes()
    context, workspace = make_context(str(source))
    token = CancellationToken()
    context._cancellation = token.view()
    operation = getattr(converter, completed)
    reached = []

    def cancel_after_operation(*args, **kwargs):
        value = operation(*args, **kwargs)
        reached.append(completed)
        next_operation.reset_mock()
        token.cancel("user_cancelled")
        return value

    next_operation = Mock(wraps=getattr(converter, following))
    monkeypatch.setattr(converter, completed, cancel_after_operation)
    monkeypatch.setattr(converter, following, next_operation)
    with pytest.raises(CancellationRequested):
        converter.MdToDocxConverter().convert(context)

    assert reached == [completed]
    next_operation.assert_not_called()
    assert workspace.registered_artifacts == []
    assert list(Path(workspace.staging_dir).iterdir()) == []
    assert source.read_bytes() == original


def test_md_to_docx_cancel_inside_link_preprocessing(monkeypatch):
    from docwen_core.links import _markdown_orchestrator, _non_embed

    source = Path(write_temp_md("# Heading\n\nPlain paragraph.\n"))
    original = source.read_bytes()
    context, workspace = make_context(str(source))
    token = CancellationToken()
    context._cancellation = token.view()
    process_links = converter.process_markdown_links
    split_blocks = _non_embed._split_fenced_code_blocks
    in_links = False
    reached = []

    def start_links(*args, **kwargs):
        nonlocal in_links
        in_links = True
        return process_links(*args, **kwargs)

    def cancel_during_scan(*args, **kwargs):
        value = split_blocks(*args, **kwargs)
        if in_links and not reached:
            reached.append(True)
            token.cancel("user_cancelled")
        return value

    image_processing = Mock(wraps=_markdown_orchestrator._replace_markdown_images)
    monkeypatch.setattr(converter, "process_markdown_links", start_links)
    monkeypatch.setattr(_non_embed, "_split_fenced_code_blocks", cancel_during_scan)
    monkeypatch.setattr(_markdown_orchestrator, "_replace_markdown_images", image_processing)
    with pytest.raises(CancellationRequested):
        converter.MdToDocxConverter().convert(context)

    assert reached == [True]
    image_processing.assert_not_called()
    assert workspace.registered_artifacts == []
    assert list(Path(workspace.staging_dir).iterdir()) == []
    assert source.read_bytes() == original
