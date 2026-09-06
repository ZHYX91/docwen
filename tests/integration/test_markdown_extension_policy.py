"""Exercise dialect switches on real DOCX bytes, independently in both directions."""

from pathlib import Path
from zipfile import ZipFile

import pytest

from docwen_core.markdown_extensions import EXTENSION_NAMES, MarkdownExtensions, resolve_markdown_extensions
from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_plugin_document.to_markdown.converter import DocxToMarkdownConverter
from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter
from docwen_runtime.config.document_styles import build_document_style_catalog
from tests.support.cancellation import FakeCancellationTokenView
from tests.support.config import FakeConfigView
from tests.support.execution import FakeExecutionContext
from tests.support.logging import FakePluginLogger
from tests.support.progress import FakeProgressSink
from tests.support.workspace import FakeWorkspaceHandle

pytestmark = pytest.mark.integration
ROOT = Path(__file__).resolve().parents[2]
DIALECTS = [
    MarkdownExtensions(),
    MarkdownExtensions.obsidian(),
    *(MarkdownExtensions(**{name: True}) for name in EXTENSION_NAMES),
]
SOURCE = """# Example

**Strong** and *emphasis*.

Table: Metrics ^metrics

| Key | Value |
| --- | --- |
| A | 1 |
| B | ^ |

See @[[#Table: Metrics]].

####### Deep

Note[^endnote:n].

[^endnote:n]: Endnote body.
"""


def _context(tmp_path: Path, source: Path, target: str, extensions: dict) -> FakeExecutionContext:
    staging = tmp_path / f"staging-{target}"
    staging.mkdir()
    ref = FileRef(path=str(source), format="markdown" if target == "docx" else "docx", category="document")
    request = ConversionRequest(
        request_id=f"dialect-{target}",
        input_refs=[ref],
        target_format=target,
        options={"markdown_extensions": extensions, "heading_merge_mode": "never", "remove_numbering": False},
        output_policy=OutputPolicy(),
    )
    return FakeExecutionContext(
        request,
        FakeWorkspaceHandle(str(source), str(staging), (ref,)),
        FakeConfigView(),
        FakeProgressSink(),
        FakeCancellationTokenView(),
        FakePluginLogger(),
        document_style_catalog=build_document_style_catalog(
            {"gui": {"language": {"locale": "zh_CN"}}},
            locales_dir=ROOT / "i18n" / "locales",
        ),
    )


@pytest.mark.parametrize("dialect", DIALECTS)
def test_input_switches_control_native_structures_and_preserve_disabled_syntax(
    tmp_path: Path, dialect: MarkdownExtensions
) -> None:
    source = tmp_path / "source.md"
    source.write_text(SOURCE, encoding="utf-8")
    context = _context(tmp_path, source, "docx", {"input": dialect.to_dict()})
    result = MdToDocxConverter().convert(context)
    assert result.success, result.error
    assert len(result.artifacts) == 1
    output = Path(result.artifacts[0].staging_path)
    with ZipFile(output) as package:
        xml = package.read("word/document.xml").decode()
        assert ("SEQ Table" in xml) is dialect.captions_references
        assert ("<w:vMerge" in xml) is dialect.structural_tables
        assert ("<w:endnoteReference" in xml) is dialect.typed_endnotes
        assert ("<w:footnoteReference" in xml) is not dialect.typed_endnotes
        if not dialect.captions_references:
            assert "Table: Metrics ^metrics" in xml
            assert "@[[#Table: Metrics]]" in xml
        assert ("####### Deep" in xml) is not dialect.extended_headings
    assert source.read_text(encoding="utf-8") == SOURCE


@pytest.mark.parametrize("dialect", DIALECTS)
def test_output_switches_are_independent_and_need_only_docx(tmp_path: Path, dialect: MarkdownExtensions) -> None:
    source = tmp_path / "source.md"
    source.write_text(SOURCE, encoding="utf-8")
    forward = MdToDocxConverter().convert(
        _context(tmp_path, source, "docx", {"input": MarkdownExtensions.obsidian().to_dict()})
    )
    assert forward.success, forward.error
    isolated = tmp_path / "input.docx"
    isolated.write_bytes(Path(forward.artifacts[0].staging_path).read_bytes())
    source.unlink()
    context = _context(tmp_path, isolated, "md", {"output": dialect.to_dict()})
    result = DocxToMarkdownConverter().convert(context)
    assert result.success, result.error
    markdown = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
    assert ("@[[#Table: Metrics]]" in markdown) is dialect.captions_references
    assert ("[^endnote:1]" in markdown) is dialect.typed_endnotes
    assert ("####### Deep" in markdown) is dialect.extended_headings
    assert "Endnote body." in markdown
    assert "| A | 1 |" in markdown
    assert "**Strong** and *emphasis*." in markdown
    assert ("| B | ^ |" in markdown) is dialect.structural_tables
    if not dialect.extended_headings:
        assert "###### Deep" in markdown
    if not dialect.typed_endnotes:
        assert "[^endnote-1]" in markdown
    if not dialect.captions_references:
        assert "^metrics" not in markdown
    if dialect != MarkdownExtensions.obsidian():
        assert any("flattened" in (item.code or "") for item in result.diagnostics)


def test_request_override_does_not_change_other_direction_or_config() -> None:
    config = FakeConfigView({"conversion": {"markdown_extensions": {"input": MarkdownExtensions.obsidian().to_dict()}}})
    options = {"markdown_extensions": {"input": {"structural_tables": False}}}
    assert not resolve_markdown_extensions(options, config, direction="input").structural_tables
    assert resolve_markdown_extensions(options, config, direction="input").captions_references
    assert resolve_markdown_extensions(options, config, direction="output") == MarkdownExtensions()
    assert resolve_markdown_extensions({}, config, direction="input") == MarkdownExtensions.obsidian()


@pytest.mark.parametrize("dialect", [MarkdownExtensions(), MarkdownExtensions.obsidian()])
def test_real_word_saved_docx_uses_current_body_and_semantics(tmp_path: Path, dialect: MarkdownExtensions) -> None:
    isolated = tmp_path / "edited.docx"
    isolated.write_bytes((ROOT / "tests/fixtures/word-save/edited-markdown-extensions.docx").read_bytes())
    result = DocxToMarkdownConverter().convert(_context(tmp_path, isolated, "md", {"output": dialect.to_dict()}))
    assert result.success, result.error
    markdown = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
    assert "Word 实测追加：从独立 DOCX 的当前内容回转。" in markdown
    assert "**粗体正文**与*斜体正文*应保留。" in markdown
    assert "| 甲 | 12 |" in markdown
    assert ("| 乙 | ^ |" in markdown) is dialect.structural_tables
    assert ("@[[#Table: 指标]]" in markdown) is dialect.captions_references
    assert ("####### 深层标题" in markdown) is dialect.extended_headings
    assert ("[^endnote:1]" in markdown) is dialect.typed_endnotes
    assert "这是尾注正文。" in markdown


def test_word_edit_that_adds_a_block_inside_caption_control_reports_invalid_structure(tmp_path: Path) -> None:
    isolated = tmp_path / "expanded-caption.docx"
    isolated.write_bytes((ROOT / "tests/fixtures/word-save/expanded-caption-control.docx").read_bytes())
    result = DocxToMarkdownConverter().convert(
        _context(tmp_path, isolated, "md", {"output": MarkdownExtensions.obsidian().to_dict()})
    )
    assert not result.success
    assert result.error is not None
    assert "one caption and one logical object" in result.error.message
    assert not result.artifacts
