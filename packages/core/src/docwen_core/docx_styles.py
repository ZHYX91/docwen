"""Immutable DOCX style identities and request-owned locale data.

The Core package owns the public semantic identity contract.  Runtime code
loads locale resources into :class:`DocumentStyleCatalog`; renderers consume
that frozen catalog without importing Runtime or reading process-global i18n
state.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

DocumentStyleKind = Literal["paragraph", "character", "table"]

SHIPPED_STYLE_LOCALES: tuple[str, ...] = (
    "zh_CN",
    "en_US",
    "de_DE",
    "es_ES",
    "fr_FR",
    "ja_JP",
    "ko_KR",
    "pt_BR",
    "ru_RU",
    "vi_VN",
    "zh_TW",
)


@dataclass(frozen=True, slots=True)
class DocumentStyleDefinition:
    """One stable semantic style identity."""

    semantic_key: str
    style_id: str
    kind: DocumentStyleKind
    based_on: str
    canonical_name: str | None = None
    locale_name_key: str | None = None

    @property
    def is_builtin(self) -> bool:
        return self.canonical_name is not None


@dataclass(frozen=True, slots=True)
class DocumentStyleFormat:
    """Normalized locale-owned defaults for a newly created style."""

    east_asia_font: str
    ascii_font: str
    font_size_pt: float
    first_line_indent_chars: int
    first_line_indent_cm: float
    spacing_after_twip: int
    spacing_before_twip: int
    bold: bool
    justification: str


@dataclass(frozen=True, slots=True)
class DocumentStyleCatalog:
    """Exact request locale plus recognition-only names from shipped locales."""

    locale: str
    output_names: tuple[tuple[str, str], ...]
    recognition_names: tuple[tuple[str, tuple[str, ...]], ...]
    formats: tuple[tuple[str, DocumentStyleFormat], ...]

    def name_for(self, semantic_key: str) -> str:
        for key, value in self.output_names:
            if key == semantic_key:
                return value
        raise KeyError(semantic_key)

    def recognition_names_for(self, semantic_key: str) -> tuple[str, ...]:
        for key, values in self.recognition_names:
            if key == semantic_key:
                return values
        raise KeyError(semantic_key)

    def format_for(self, semantic_key: str) -> DocumentStyleFormat | None:
        for key, value in self.formats:
            if key == semantic_key:
                return value
        return None


def _builtin(
    semantic_key: str,
    style_id: str,
    kind: DocumentStyleKind,
    canonical_name: str,
    based_on: str,
) -> DocumentStyleDefinition:
    return DocumentStyleDefinition(
        semantic_key=semantic_key,
        style_id=style_id,
        kind=kind,
        based_on=based_on,
        canonical_name=canonical_name,
    )


def _custom(
    semantic_key: str,
    style_id: str,
    kind: DocumentStyleKind,
    based_on: str,
) -> DocumentStyleDefinition:
    return DocumentStyleDefinition(
        semantic_key=semantic_key,
        style_id=style_id,
        kind=kind,
        based_on=based_on,
        locale_name_key=semantic_key,
    )


BUILTIN_DOCUMENT_STYLES: tuple[DocumentStyleDefinition, ...] = (
    *(
        _builtin(
            f"heading_{level}",
            f"Heading{level}",
            "paragraph",
            f"heading {level}",
            "Normal",
        )
        for level in range(1, 10)
    ),
    _builtin("footnote_text", "FootnoteText", "paragraph", "footnote text", "Normal"),
    _builtin(
        "footnote_reference",
        "FootnoteReference",
        "character",
        "footnote reference",
        "DefaultParagraphFont",
    ),
    _builtin("endnote_text", "EndnoteText", "paragraph", "endnote text", "Normal"),
    _builtin(
        "endnote_reference",
        "EndnoteReference",
        "character",
        "endnote reference",
        "DefaultParagraphFont",
    ),
    _builtin("caption", "Caption", "paragraph", "caption", "Normal"),
    _builtin("bibliography", "Bibliography", "paragraph", "Bibliography", "Normal"),
    _builtin("hyperlink", "Hyperlink", "character", "Hyperlink", "DefaultParagraphFont"),
)

CUSTOM_DOCUMENT_STYLES: tuple[DocumentStyleDefinition, ...] = (
    _custom("body_paragraph", "DocWenBodyParagraph", "paragraph", "Normal"),
    _custom("image_paragraph", "DocWenImageParagraph", "paragraph", "Normal"),
    _custom("code_block", "DocWenCodeBlock", "paragraph", "Normal"),
    _custom("inline_code", "DocWenInlineCode", "character", "DefaultParagraphFont"),
    _custom("formula_block", "DocWenFormulaBlock", "paragraph", "Normal"),
    _custom("inline_formula", "DocWenInlineFormula", "character", "DefaultParagraphFont"),
    _custom("list_block", "DocWenListBlock", "paragraph", "Normal"),
    _custom("horizontal_rule_1", "DocWenHorizontalRule1", "paragraph", "Normal"),
    _custom("horizontal_rule_2", "DocWenHorizontalRule2", "paragraph", "Normal"),
    _custom("horizontal_rule_3", "DocWenHorizontalRule3", "paragraph", "Normal"),
    _custom("table_content", "DocWenTableContent", "paragraph", "Normal"),
    _custom("table_header", "DocWenTableHeader", "paragraph", "Normal"),
    _custom("three_line_table", "DocWenThreeLineTable", "table", "TableNormal"),
    _custom("table_grid", "DocWenTableGrid", "table", "TableNormal"),
    *(_custom(f"quote_{level}", f"DocWenQuote{level}", "paragraph", "Normal") for level in range(1, 10)),
    _custom("figure_caption", "DocWenFigureCaption", "paragraph", "Caption"),
    _custom("table_caption", "DocWenTableCaption", "paragraph", "Caption"),
    _custom("equation_caption", "DocWenEquationCaption", "paragraph", "Caption"),
    _custom("code_block_caption", "DocWenCodeBlockCaption", "paragraph", "Caption"),
)

MANAGED_DOCUMENT_STYLES: tuple[DocumentStyleDefinition, ...] = (
    *BUILTIN_DOCUMENT_STYLES,
    *CUSTOM_DOCUMENT_STYLES,
)

CUSTOM_DOCUMENT_STYLE_KEYS: tuple[str, ...] = tuple(definition.semantic_key for definition in CUSTOM_DOCUMENT_STYLES)


__all__ = [
    "BUILTIN_DOCUMENT_STYLES",
    "CUSTOM_DOCUMENT_STYLES",
    "CUSTOM_DOCUMENT_STYLE_KEYS",
    "MANAGED_DOCUMENT_STYLES",
    "SHIPPED_STYLE_LOCALES",
    "DocumentStyleCatalog",
    "DocumentStyleDefinition",
    "DocumentStyleFormat",
    "DocumentStyleKind",
]
