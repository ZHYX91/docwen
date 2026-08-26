from __future__ import annotations

import pytest

pytestmark = pytest.mark.unit


class _Style:
    def __init__(self, name: str) -> None:
        self.name = name


class _Paragraph:
    def __init__(self, style_name: str) -> None:
        self.style = _Style(style_name)


def test_style_detector_config_from_document_config_consumes_code_and_quote_aliases() -> None:
    from docwen_core.docx_parsing.format_features import (
        detect_paragraph_style_type,
        style_detector_config_from_document_config,
    )

    config = style_detector_config_from_document_config(
        {
            "style": {
                "code": {
                    "docx_to_md": {
                        "paragraph_style_aliases": ["Console Listing", "console listing"],
                        "character_style_aliases": ["Console Inline", "console inline"],
                        "full_paragraph_as_block": False,
                        "fuzzy_match_enabled": False,
                        "shading": {"wps_enabled": False, "word_enabled": True},
                    }
                },
                "quote": {
                    "docx_to_md": {
                        "level_style_aliases": {"Pull Quote": 3, "pull quote": 8},
                        "paragraph_style_aliases": ["Aside", "aside"],
                        "character_style_aliases": ["Aside Char", "aside char"],
                        "full_paragraph_as_block": False,
                        "fuzzy_match_enabled": False,
                    }
                },
            }
        }
    )

    assert config is not None
    assert config.code_character_style_names == frozenset({"Console Inline"})
    assert config.quote_character_style_names == frozenset({"Aside Char"})
    assert config.code_full_paragraph_as_block is False
    assert config.quote_full_paragraph_as_block is False
    assert config.wps_shading_enabled is False
    assert config.word_shading_enabled is True
    assert detect_paragraph_style_type(_Paragraph("Console Listing"), config=config) == ("code_block", True)
    assert detect_paragraph_style_type(_Paragraph("Aside"), config=config) == ("quote", 1)
    assert detect_paragraph_style_type(_Paragraph("Pull Quote"), config=config) == ("quote", 3)
    assert detect_paragraph_style_type(_Paragraph("Console Listing Notes"), config=config) == (None, None)
    assert detect_paragraph_style_type(_Paragraph("Aside Notes"), config=config) == (None, None)
    assert detect_paragraph_style_type(_Paragraph("Pull Quote Notes"), config=config) == (None, None)
    assert detect_paragraph_style_type(_Paragraph("Decoder Notes"), config=config) == (None, None)


def test_configured_quote_alias_digit_overrides_default_level() -> None:
    from docwen_core.docx_parsing.format_features import StyleDetectorConfig, detect_paragraph_style_type

    config = StyleDetectorConfig(quote_style_patterns=(("Custom Quote", 2),))

    assert detect_paragraph_style_type(_Paragraph("Custom Quote"), config=config) == ("quote", 2)
    assert detect_paragraph_style_type(_Paragraph("Custom Quote 4"), config=config) == ("quote", 4)


def test_bundled_quote_patterns_do_not_match_inside_ascii_words() -> None:
    from docwen_core.docx_parsing.format_features import detect_paragraph_style_type

    assert detect_paragraph_style_type(_Paragraph("Quote 2")) == ("quote", 2)
    assert detect_paragraph_style_type(_Paragraph("Misquote 2")) == (None, None)
    assert detect_paragraph_style_type(_Paragraph("Unblockquote 3")) == (None, None)


def test_configured_fuzzy_keyword_matches_only_at_style_word_boundary() -> None:
    from docwen_core.docx_parsing.format_features import StyleDetectorConfig, detect_paragraph_style_type

    config = StyleDetectorConfig(code_block_style_fragments=("snippet",))

    assert detect_paragraph_style_type(_Paragraph("Acme Snippet Notes"), config=config) == ("code_block", True)
    assert detect_paragraph_style_type(_Paragraph("AcmeSnippetNotes"), config=config) == (None, None)
    assert detect_paragraph_style_type(_Paragraph("Decoder Notes"), config=config) == (None, None)


@pytest.mark.parametrize(
    ("style_name", "expected"),
    [
        ("Code", ("code_block", True)),
        ("Source", ("code_block", True)),
        ("Programming", ("code_block", True)),
        ("Program", ("code_block", True)),
        ("代码", ("code_block", True)),
        ("程序", ("code_block", True)),
        ("Quote", ("quote", 1)),
        ("Blockquote", ("quote", 1)),
        ("引用", ("quote", 1)),
    ],
)
def test_legacy_proofread_style_names_remain_exact_builtins(style_name: str, expected: tuple[str, object]) -> None:
    from docwen_core.docx_parsing.format_features import StyleDetectorConfig, detect_paragraph_style_type

    exact_only = StyleDetectorConfig(code_fuzzy_match_enabled=False, quote_fuzzy_match_enabled=False)
    assert detect_paragraph_style_type(_Paragraph(style_name.upper()), config=exact_only) == expected
    assert detect_paragraph_style_type(_Paragraph(f"{style_name} Notes"), config=exact_only) == (None, None)


def test_configured_quote_fragment_obeys_ascii_word_boundaries() -> None:
    from docwen_core.docx_parsing.format_features import StyleDetectorConfig, detect_paragraph_style_type

    config = StyleDetectorConfig(quote_style_patterns=(("Direct Quote Alias", 1),))

    assert detect_paragraph_style_type(_Paragraph("My Direct Quote Alias Style"), config=config) == ("quote", 1)
    assert detect_paragraph_style_type(_Paragraph("Indirect Quote Alias"), config=config) == (None, None)
    assert detect_paragraph_style_type(_Paragraph("Direct Quote Aliased"), config=config) == (None, None)


def test_code_block_accumulator_indents_every_physical_line_in_list_context() -> None:
    from docwen_core.docx_parsing.format_features import CodeBlockAccumulator

    accumulator = CodeBlockAccumulator(indent_spaces=4)
    accumulator.start(list_level=1)
    accumulator.add_line("One\nTwo")

    assert accumulator.finalize() == "    ```\n    One\n    Two\n    ```"


@pytest.mark.parametrize(
    "style_name",
    [
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
    ],
)
def test_bundled_localized_code_block_style_names_are_detected(style_name: str) -> None:
    from docwen_core.docx_parsing.format_features import detect_paragraph_style_type

    assert detect_paragraph_style_type(_Paragraph(style_name.upper())) == ("code_block", True)


@pytest.mark.parametrize(
    ("style_name", "level"),
    [
        ("Zitat 1", 1),
        ("Quote 2", 2),
        ("Cita 3", 3),
        ("Citation 4", 4),
        ("引用 5", 5),
        ("인용 6", 6),
        ("Citação 7", 7),
        ("Цитата 8", 8),
        ("Trích dẫn 9", 9),
    ],
)
def test_bundled_localized_quote_style_names_preserve_level(style_name: str, level: int) -> None:
    from docwen_core.docx_parsing.format_features import detect_paragraph_style_type

    assert detect_paragraph_style_type(_Paragraph(style_name.upper())) == ("quote", level)


def test_docx_markdown_formatting_config_from_conversion_config() -> None:
    from docwen_core.docx_parsing.format_features import docx_markdown_formatting_config_from_conversion_config

    config = docx_markdown_formatting_config_from_conversion_config(
        {
            "docx_to_md": {
                "preserve_formatting": False,
                "preserve_heading_formatting": True,
                "preserve_table_header_formatting": True,
            }
        }
    )

    assert config.preserve_formatting is False
    assert config.preserve_heading_formatting is True
    assert config.preserve_table_header_formatting is True


def test_docx_markdown_syntax_config_from_conversion_config() -> None:
    from docwen_core.docx_parsing.format_features import docx_markdown_syntax_config_from_conversion_config

    config = docx_markdown_syntax_config_from_conversion_config(
        {
            "syntax": {
                "bold": "underscore",
                "italic": "underscore",
                "strikethrough": "html",
                "highlight": "html",
                "superscript": "extended",
                "subscript": "extended",
                "unordered_list": "plus",
                "indent_spaces": 2,
            }
        }
    )

    assert config.bold == "underscore"
    assert config.italic == "underscore"
    assert config.strikethrough == "html"
    assert config.highlight == "html"
    assert config.superscript == "extended"
    assert config.subscript == "extended"
    assert config.unordered_list == "plus"
    assert config.indent_spaces == 2


def test_docx_markdown_syntax_config_rejects_unknown_values() -> None:
    from docwen_core.docx_parsing.format_features import docx_markdown_syntax_config_from_conversion_config

    config = docx_markdown_syntax_config_from_conversion_config(
        {
            "syntax": {
                "bold": "stars",
                "italic": "slashes",
                "strikethrough": "gfm",
                "highlight": "mark",
                "superscript": "caret",
                "subscript": "tilde",
                "unordered_list": "dot",
                "indent_spaces": 3,
            }
        }
    )

    assert config.bold == "asterisk"
    assert config.italic == "asterisk"
    assert config.strikethrough == "extended"
    assert config.highlight == "extended"
    assert config.superscript == "html"
    assert config.subscript == "html"
    assert config.unordered_list == "dash"
    assert config.indent_spaces == 4
