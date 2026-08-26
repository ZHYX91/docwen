"""Pure unit tests for proofread paragraph skip policy."""

from __future__ import annotations

from types import SimpleNamespace

import pytest

from docwen_plugin_proofread.skip_policy import (
    ProofreadSkipOptions,
    resolve_skip_options,
    should_skip_docx_paragraph,
)

pytestmark = pytest.mark.unit


class _Style:
    def __init__(self, name: str) -> None:
        self.name = name


class _Font:
    def __init__(self, name: str | None) -> None:
        self.name = name


class _Run:
    def __init__(self, font_name: str | None = None) -> None:
        self.font = _Font(font_name)


class _Paragraph:
    def __init__(
        self,
        text: str,
        *,
        style_name: str | None = None,
        font_names: tuple[str | None, ...] = (),
    ) -> None:
        self.text = text
        self.style = _Style(style_name) if style_name is not None else None
        self.runs = [_Run(font_name) for font_name in font_names]
        self._element = SimpleNamespace(xml="<w:p/>")


def _context(config: dict | None = None, options: dict | None = None):
    request = SimpleNamespace(options=options or {})
    return SimpleNamespace(config=config or {}, request=request)


def test_code_style_paragraph_is_skipped_when_enabled() -> None:
    para = _Paragraph("print('hello')", style_name="Code Block")
    opts = ProofreadSkipOptions(code_blocks=True, quote_blocks=False)

    assert should_skip_docx_paragraph(para, opts) is True


def test_code_style_paragraph_is_not_skipped_when_disabled() -> None:
    para = _Paragraph("print('hello')", style_name="Code Block")
    opts = ProofreadSkipOptions(code_blocks=False, quote_blocks=False)

    assert should_skip_docx_paragraph(para, opts) is False


def test_chinese_code_style_paragraph_is_skipped() -> None:
    para = _Paragraph("示例代码", style_name="代码块")
    opts = ProofreadSkipOptions(code_blocks=True, quote_blocks=False)

    assert should_skip_docx_paragraph(para, opts) is True


@pytest.mark.parametrize(
    "style_name",
    (
        "Codeblock",
        "Code Block",
        "Bloque de código",
        "Bloc de code",
        "コードブロック",
        "코드 블록",
        "Bloco de código",
        "Блок кода",
        "Khối mã",
        "代码块",
        "代碼塊",
    ),
)
def test_every_bundled_locale_code_style_is_skipped(style_name: str) -> None:
    para = _Paragraph("literal text", style_name=style_name)

    assert should_skip_docx_paragraph(para, ProofreadSkipOptions()) is True


def test_custom_document_code_alias_is_skipped_from_request_snapshot() -> None:
    context = _context(
        config={
            "document": {
                "style": {
                    "code": {
                        "docx_to_md": {
                            "paragraph_style_aliases": ["AcmeSnippet"],
                            "fuzzy_match_enabled": False,
                        }
                    }
                }
            }
        }
    )
    para = _Paragraph("custom literal", style_name="AcmeSnippet")

    assert should_skip_docx_paragraph(para, resolve_skip_options(context)) is True


def test_fuzzy_disabled_code_alias_does_not_skip_suffix_or_decoder_style() -> None:
    context = _context(
        config={
            "document": {
                "style": {
                    "code": {
                        "docx_to_md": {
                            "paragraph_style_aliases": ["AcmeSnippet"],
                            "fuzzy_match_enabled": False,
                        }
                    }
                }
            }
        }
    )
    options = resolve_skip_options(context)

    assert should_skip_docx_paragraph(_Paragraph("ordinary", style_name="AcmeSnippetNotes"), options) is False
    assert should_skip_docx_paragraph(_Paragraph("ordinary", style_name="Decoder Notes"), options) is False


def test_fuzzy_disabled_quote_alias_does_not_skip_suffix_style() -> None:
    context = _context(
        config={
            "document": {
                "style": {
                    "quote": {
                        "docx_to_md": {
                            "paragraph_style_aliases": ["Aside"],
                            "fuzzy_match_enabled": False,
                        }
                    }
                }
            },
            "proofread": {"skip": {"quote_blocks": True}},
        }
    )

    assert (
        should_skip_docx_paragraph(
            _Paragraph("ordinary", style_name="Aside Notes"),
            resolve_skip_options(context),
        )
        is False
    )


@pytest.mark.parametrize(
    "style_name",
    ("BLOQUE DE CÓDIGO", "BLOCO DE CÓDIGO", "БЛОК КОДА", "KHỐI MÃ"),
)
def test_bundled_locale_code_style_matching_is_case_insensitive(style_name: str) -> None:
    assert should_skip_docx_paragraph(_Paragraph("literal", style_name=style_name), ProofreadSkipOptions()) is True


@pytest.mark.parametrize("style_name", ("CITATION 3", "CITAÇÃO 5", "ЦИТАТА 6", "TRÍCH DẪN 9"))
def test_bundled_locale_quote_style_matching_is_case_insensitive(style_name: str) -> None:
    options = ProofreadSkipOptions(code_blocks=False, quote_blocks=True)

    assert should_skip_docx_paragraph(_Paragraph("quoted", style_name=style_name), options) is True


@pytest.mark.parametrize("style_name", ("Code", "Source", "Programming", "Program", "代码", "程序"))
def test_legacy_code_style_names_are_exactly_skipped(style_name: str) -> None:
    from docwen_core.docx_parsing.format_features import StyleDetectorConfig

    options = ProofreadSkipOptions(
        code_blocks=True,
        quote_blocks=False,
        style_detector_config=StyleDetectorConfig(code_fuzzy_match_enabled=False),
    )

    assert should_skip_docx_paragraph(_Paragraph("literal", style_name=style_name), options) is True
    assert should_skip_docx_paragraph(_Paragraph("ordinary", style_name=f"{style_name} Notes"), options) is False


@pytest.mark.parametrize("style_name", ("Quote", "Blockquote", "引用"))
def test_legacy_quote_style_names_are_exactly_skipped(style_name: str) -> None:
    from docwen_core.docx_parsing.format_features import StyleDetectorConfig

    options = ProofreadSkipOptions(
        code_blocks=False,
        quote_blocks=True,
        style_detector_config=StyleDetectorConfig(quote_fuzzy_match_enabled=False),
    )

    assert should_skip_docx_paragraph(_Paragraph("quoted", style_name=style_name), options) is True
    assert should_skip_docx_paragraph(_Paragraph("ordinary", style_name=f"{style_name} Notes"), options) is False


def test_quote_style_paragraph_is_skipped_when_enabled() -> None:
    para = _Paragraph("quoted text", style_name="Block Quote")
    opts = ProofreadSkipOptions(code_blocks=False, quote_blocks=True)

    assert should_skip_docx_paragraph(para, opts) is True


def test_markdown_quote_prefix_is_skipped_when_enabled() -> None:
    para = _Paragraph("> quoted text", style_name="Normal")
    opts = ProofreadSkipOptions(code_blocks=False, quote_blocks=True)

    assert should_skip_docx_paragraph(para, opts) is True


def test_normal_paragraph_is_not_skipped() -> None:
    para = _Paragraph("normal text", style_name="Normal", font_names=("Arial",))
    opts = ProofreadSkipOptions(code_blocks=True, quote_blocks=True)

    assert should_skip_docx_paragraph(para, opts) is False


def test_mixed_text_and_drawing_paragraph_is_not_misclassified_as_code() -> None:
    para = _Paragraph("text beside image", style_name="Normal")
    para._element = SimpleNamespace(xml="<w:p><w:r><w:drawing/></w:r></w:p>")
    opts = ProofreadSkipOptions(code_blocks=True, quote_blocks=False)

    assert should_skip_docx_paragraph(para, opts) is False


def test_resolve_skip_options_reads_config_skip_section() -> None:
    context = _context(config={"proofread": {"skip": {"code_blocks": False, "quote_blocks": True}}})

    assert resolve_skip_options(context) == ProofreadSkipOptions(
        code_blocks=False,
        quote_blocks=True,
    )


def test_request_options_override_config_skip_section() -> None:
    context = _context(
        config={"proofread": {"skip": {"code_blocks": False, "quote_blocks": False}}},
        options={"skip_code_blocks": True, "skip_quote_blocks": True},
    )

    assert resolve_skip_options(context) == ProofreadSkipOptions(
        code_blocks=True,
        quote_blocks=True,
    )


def test_runtime_read_only_config_view_projects_exact_style_alias() -> None:
    from docwen_runtime._execution_context import (  # pyright: ignore[reportPrivateUsage]
        _RuntimeReadOnlyConfigView,
    )

    config = _RuntimeReadOnlyConfigView(
        {
            "document": {
                "style": {
                    "code": {
                        "docx_to_md": {
                            "paragraph_style_aliases": ["RuntimeSnippet"],
                            "fuzzy_match_enabled": False,
                        }
                    }
                }
            }
        }
    )
    context = SimpleNamespace(config=config, request=SimpleNamespace(options={}))
    options = resolve_skip_options(context)

    assert should_skip_docx_paragraph(_Paragraph("literal", style_name="RuntimeSnippet"), options) is True
    assert should_skip_docx_paragraph(_Paragraph("ordinary", style_name="RuntimeSnippetNotes"), options) is False


def test_legacy_option_names_are_not_supported() -> None:
    context = _context(
        config={"proofread": {"skip": {"code_blocks": True, "quote_blocks": False}}},
        options={"is_code_paragraph": False},
    )

    assert resolve_skip_options(context).code_blocks is True
