"""Optional generated marker output and safe reimport of marked headings."""

from types import SimpleNamespace

import pytest
from docx import Document

from docwen_application.conversion_capabilities import CAPABILITY_BY_ID
from docwen_application.conversion_contracts import DOCX_TO_MARKDOWN_CAPABILITY_ID
from docwen_application.conversion_options import resolve_conversion_options
from docwen_core.text.heading_numbering import HeadingFormatter
from docwen_plugin_document.to_markdown.converter import DocxToMarkdownConverter

from ._request_scoped_docx_policy_support import _context, _markdown_from_result
from ._to_markdown_standard_parity_support import _ExactSchemeRegistry

pytestmark = pytest.mark.contract
WJ = "\u2060"


@pytest.mark.parametrize(
    ("config_value", "override", "expected"),
    [(None, None, False), (True, None, True), (True, False, False), (False, True, True)],
)
def test_option_overrides_request_snapshot(tmp_path, config_value, override, expected):
    source = tmp_path / "source.docx"
    document = Document()
    document.add_heading("Title", level=1)
    document.save(source)
    config = {} if config_value is None else {"document": {"export_numbering_markers": config_value}}
    options = {"add_numbering": True, "numbering_scheme": "exact"}
    if override is not None:
        options["export_numbering_markers"] = override
    options = resolve_conversion_options(options, CAPABILITY_BY_ID[DOCX_TO_MARKDOWN_CAPABILITY_ID])
    if override is None:
        assert "export_numbering_markers" not in options
    context = _context(tmp_path, source, request_id="markers", config=config, options=options)
    context._numbering_registry = _ExactSchemeRegistry()
    markdown = _markdown_from_result(DocxToMarkdownConverter().convert(context))
    assert (f"# {WJ}1{WJ} Title" if expected else "# 1 Title") in markdown
    assert (WJ in markdown) is expected


@pytest.mark.parametrize("bold", [False, True])
@pytest.mark.parametrize("remove", [False, True])
def test_marked_docx_heading_preserved_or_replaced_without_duplicate(bold, remove):
    document = Document()
    heading = document.add_heading("", level=1)
    heading.add_run(f"{WJ}9.{WJ} ").bold = bold
    heading.add_run("Title").bold = bold
    converter = DocxToMarkdownConverter()
    converter._preserve_heading_formatting = True
    formatter = HeadingFormatter({"level_1": {"format": "{1.arabic_half}. "}}, mark_generated_numbers=True)
    lines, _ = converter._process_paragraph(
        heading._element, {id(heading._element): heading}, remove_numbering=remove, heading_formatter=formatter
    )
    output = "\n".join(lines)
    assert output.count(WJ) == 2
    assert f"{WJ}{'1' if remove else '9'}.{WJ}" in output
    assert "Title" in output


def test_authenticated_heading_bypasses_historical_marker_output(monkeypatch):
    document = Document()
    heading = document.add_heading("Title", level=1)
    converter = DocxToMarkdownConverter()
    monkeypatch.setattr(
        converter,
        "_resolved_v4_recovery",
        SimpleNamespace(heading_import=lambda _para: SimpleNamespace(diagnostics=())),
    )
    formatter = HeadingFormatter({"level_1": {"format": "{1.arabic_half}. "}}, mark_generated_numbers=True)
    lines, _ = converter._process_paragraph(
        heading._element, {id(heading._element): heading}, remove_numbering=True, heading_formatter=formatter
    )
    assert "# Title" in lines
    assert WJ not in "\n".join(lines)


def test_markers_setting_does_not_number_without_add(tmp_path):
    source = tmp_path / "plain.docx"
    document = Document()
    document.add_heading("Title", level=1)
    document.save(source)
    context = _context(
        tmp_path,
        source,
        request_id="plain",
        config={"document": {"export_numbering_markers": True}},
        options={"add_numbering": False},
    )
    markdown = _markdown_from_result(DocxToMarkdownConverter().convert(context))
    assert "# Title" in markdown
    assert WJ not in markdown


def test_formatted_incomplete_marker_remains_unchanged():
    document = Document()
    heading = document.add_heading("", level=1)
    heading.add_run(f"1. Title{WJ}").bold = True
    converter = DocxToMarkdownConverter()
    converter._preserve_heading_formatting = True
    formatter = HeadingFormatter({"level_1": {"format": "{1.arabic_half}. "}}, mark_generated_numbers=True)
    lines, _ = converter._process_paragraph(
        heading._element, {id(heading._element): heading}, remove_numbering=True, heading_formatter=formatter
    )
    output = "\n".join(lines)
    assert output.count(WJ) == 1
    assert f"1\\. Title{WJ}" in output
