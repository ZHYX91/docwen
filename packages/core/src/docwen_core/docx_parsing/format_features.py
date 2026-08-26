"""Generic paragraph format extraction helpers."""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from collections.abc import Iterator, Mapping
from dataclasses import dataclass
from typing import Any

from docwen_core.docx_parsing.xml_ns import NS_W

_CJK_TEXT_RE = re.compile(
    r"[\u3040-\u30ff\u3100-\u312f\u3130-\u318f\u31f0-\u31ff"
    r"\u3400-\u4dbf\u4e00-\u9fff\uac00-\ud7af\uf900-\ufaff]"
)
_JAPANESE_TEXT_RE = re.compile(r"[\u3040-\u30ff\u31f0-\u31ff]")
_KOREAN_TEXT_RE = re.compile(r"[\u3130-\u318f\uac00-\ud7af]")
_BOPOMOFO_TEXT_RE = re.compile(r"[\u3100-\u312f]")
_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _representative_run(para: Any) -> Any | None:
    """Choose the run that best represents a paragraph's visible text."""

    runs = list(getattr(para, "runs", ()) or ())
    if not runs:
        return None
    for run in runs:
        if _CJK_TEXT_RE.search(str(getattr(run, "text", "") or "")):
            return run
    for run in runs:
        if str(getattr(run, "text", "") or "").strip():
            return run
    return runs[0]


def _iter_style_chain(style: Any) -> Iterator[Any]:
    """Yield a style and its bases without trusting producer-owned cycles."""

    seen: set[int] = set()
    current = style
    while current is not None and id(current) not in seen:
        seen.add(id(current))
        yield current
        current = getattr(current, "base_style", None)


def _style_rpr(style: Any) -> Any | None:
    element = getattr(style, "element", None)
    if element is None:
        element = getattr(style, "_element", None)
    if element is None:
        return None
    r_pr = getattr(element, "rPr", None)
    return r_pr if r_pr is not None else element.find(f"{{{NS_W}}}rPr")


def _rfonts(r_pr: Any) -> Any | None:
    if r_pr is None:
        return None
    try:
        return r_pr.find(f"{{{NS_W}}}rFonts")
    except Exception:
        return None


def _rfonts_value(r_pr: Any, slot: str) -> str | None:
    r_fonts = _rfonts(r_pr)
    if r_fonts is None:
        return None
    value = r_fonts.get(f"{{{NS_W}}}{slot}")
    return str(value) if value else None


def _east_asia_script(language: str | None, text: str) -> str:
    normalized = (language or "").strip().lower().replace("_", "-")
    if normalized.startswith("ja"):
        return "Jpan"
    if normalized.startswith("ko"):
        return "Hang"
    if normalized.startswith("zh"):
        if "hant" in normalized or any(region in normalized.split("-") for region in ("tw", "hk", "mo")):
            return "Hant"
        return "Hans"
    if _JAPANESE_TEXT_RE.search(text):
        return "Jpan"
    if _KOREAN_TEXT_RE.search(text):
        return "Hang"
    if _BOPOMOFO_TEXT_RE.search(text):
        return "Hant"
    return "Hans"


def _east_asia_language(para: Any, r_pr_sources: list[Any]) -> str | None:
    for r_pr in r_pr_sources:
        if r_pr is None:
            continue
        try:
            lang = r_pr.find(f"{{{NS_W}}}lang")
            if lang is not None:
                value = lang.get(f"{{{NS_W}}}eastAsia") or lang.get(f"{{{NS_W}}}val")
                if value:
                    return str(value)
        except Exception:
            continue
    try:
        settings = para.part.document.settings.element
        theme_font_lang = settings.find(f"{{{NS_W}}}themeFontLang")
        if theme_font_lang is not None:
            value = theme_font_lang.get(f"{{{NS_W}}}eastAsia") or theme_font_lang.get(f"{{{NS_W}}}val")
            if value:
                return str(value)
    except Exception:
        pass
    return None


def _resolve_theme_font(para: Any, token: str, *, script: str | None) -> str | None:
    token_lower = token.lower()
    if token_lower.startswith("major"):
        family_name = "majorFont"
    elif token_lower.startswith("minor"):
        family_name = "minorFont"
    else:
        return None
    try:
        relationships = para.part.document.part.rels.values()
        theme_part = next(rel.target_part for rel in relationships if str(rel.reltype).endswith("/theme"))
        root = ET.fromstring(theme_part.blob)
        family = root.find(f".//{{{_NS_A}}}{family_name}")
        if family is None:
            return None
        if script is not None:
            for font in family.findall(f"{{{_NS_A}}}font"):
                if font.get("script") == script and font.get("typeface"):
                    return str(font.get("typeface"))
            fallback = family.find(f"{{{_NS_A}}}ea")
        else:
            fallback = family.find(f"{{{_NS_A}}}latin")
        value = fallback.get("typeface") if fallback is not None else None
        return str(value) if value else None
    except Exception:
        return None


def _effective_font_name(para: Any, run: Any, r_pr_sources: list[Any]) -> str | None:
    text = str(getattr(run, "text", "") or "")
    prefer_east_asia = bool(_CJK_TEXT_RE.search(text))
    if prefer_east_asia:
        language = _east_asia_language(para, r_pr_sources)
        script = _east_asia_script(language, text)
        for r_pr in r_pr_sources:
            concrete = _rfonts_value(r_pr, "eastAsia")
            if concrete:
                return concrete
            theme = _rfonts_value(r_pr, "eastAsiaTheme")
            if theme:
                resolved = _resolve_theme_font(para, theme, script=script)
                if resolved:
                    return resolved
        # Latin slots are a last-resort degradation only after every effective
        # East Asian source has been exhausted.
        for r_pr in r_pr_sources:
            for slot in ("hAnsi", "ascii", "cs"):
                concrete = _rfonts_value(r_pr, slot)
                if concrete:
                    return concrete
        return None

    for r_pr in r_pr_sources:
        for slot in ("ascii", "hAnsi", "cs"):
            concrete = _rfonts_value(r_pr, slot)
            if concrete:
                return concrete
        for slot in ("asciiTheme", "hAnsiTheme", "cstheme"):
            theme = _rfonts_value(r_pr, slot)
            if theme:
                resolved = _resolve_theme_font(para, theme, script=None)
                if resolved:
                    return resolved
    for r_pr in r_pr_sources:
        concrete = _rfonts_value(r_pr, "eastAsia")
        if concrete:
            return concrete
    return None


def _font_size_from_rpr(r_pr: Any) -> float | None:
    if r_pr is None:
        return None
    try:
        size = r_pr.find(f"{{{NS_W}}}sz")
        if size is None:
            size = r_pr.find(f"{{{NS_W}}}szCs")
        value = size.get(f"{{{NS_W}}}val") if size is not None else None
        return float(value) / 2.0 if value is not None else None
    except (AttributeError, TypeError, ValueError):
        return None


def _doc_defaults_rpr(para: Any) -> Any | None:
    try:
        styles = para.part.document.styles.element
        doc_defaults = styles.find(f"{{{NS_W}}}docDefaults")
        if doc_defaults is None:
            return None
        r_pr_default = doc_defaults.find(f"{{{NS_W}}}rPrDefault")
        return r_pr_default.find(f"{{{NS_W}}}rPr") if r_pr_default is not None else None
    except Exception:
        return None


def extract_font_info(para: Any) -> tuple[str, float | None]:
    """Return the effective representative font name and size.

    Direct formatting on the first CJK-bearing run wins. Missing name and size
    components are then resolved independently through the run style,
    paragraph style/base styles, Normal, and document defaults. This avoids a
    decorative leading run hiding the format that Gongwen recognition needs.
    """

    run = _representative_run(para)
    if run is None:
        return "", None

    r_pr_sources: list[Any] = [getattr(getattr(run, "_r", None), "rPr", None)]
    r_pr_sources.extend(_style_rpr(style) for style in _iter_style_chain(getattr(run, "style", None)))

    paragraph_styles = list(_iter_style_chain(getattr(para, "style", None)))
    r_pr_sources.extend(_style_rpr(style) for style in paragraph_styles)
    try:
        normal_style = para.part.document.styles["Normal"]
    except Exception:
        normal_style = None
    normal_style_id = getattr(normal_style, "style_id", None)
    if normal_style is not None and all(
        getattr(style, "style_id", None) != normal_style_id for style in paragraph_styles
    ):
        r_pr_sources.extend(_style_rpr(style) for style in _iter_style_chain(normal_style))
    r_pr_sources.append(_doc_defaults_rpr(para))

    name = _effective_font_name(para, run, r_pr_sources)
    size: float | None = None
    for r_pr in r_pr_sources:
        if size is None:
            size = _font_size_from_rpr(r_pr)
        if size is not None:
            break
    return name or "", size


def extract_alignment(para: Any) -> str:
    alignment = getattr(para, "alignment", None)
    return alignment.name if alignment is not None else ""


def extract_outline_level(para: Any) -> int | None:
    try:
        p_pr = para._p.pPr
        if p_pr is not None:
            outline = p_pr.find(f".//{{{NS_W}}}outlineLvl")
            if outline is not None:
                val = outline.get(f"{{{NS_W}}}val")
                if val is not None:
                    level = int(val)
                    # OOXML uses 0..8 for outline levels 1..9 and 9 as
                    # Word's explicit "Body Text" sentinel.  Treating that
                    # sentinel as a heading turns ordinary producer-authored
                    # body paragraphs into a fictitious tenth heading level.
                    return level if 0 <= level <= 8 else None
    except Exception:
        pass

    style = getattr(para, "style", None)
    style_name = getattr(style, "name", "") or ""
    match = re.search(r"Heading\s+([1-9])", style_name, re.IGNORECASE)
    if match:
        return int(match.group(1)) - 1
    return None


# ── Paragraph style type detection ────────────────────────────────────────

# Well‑known code block style name fragments (case-insensitive match).
_CODE_BLOCK_STYLE_FRAGMENTS = [
    "code block",
    "codeblock",
    "代码块",
    "源代码",
    "code",
    "source code",
    "program",
    "listing",
    "html code",
    "xml code",
    "json code",
]

# Well‑known quote style name fragments, with optional level digit.
_QUOTE_STYLE_PATTERNS = [
    (re.compile(r"(?<![0-9A-Za-z])quote\s*(\d)(?![0-9A-Za-z])", re.IGNORECASE), lambda m: int(m.group(1))),
    (re.compile(r"引用\s*(\d)"), lambda m: int(m.group(1))),
    (
        re.compile(r"(?<![0-9A-Za-z])block\s*quote\s*(\d)(?![0-9A-Za-z])", re.IGNORECASE),
        lambda m: int(m.group(1)),
    ),
    (re.compile(r"(?<![0-9A-Za-z])quotation\s*(\d)(?![0-9A-Za-z])", re.IGNORECASE), lambda m: int(m.group(1))),
]
_QUOTE_GENERIC_NAMES = {
    "quote",
    "引用",
    "quotation",
    "block quote",
    "block text",
    "blocktext",
    "intense quote",
    "intensequote",
}

# Exact names emitted by every bundled locale template.  Keep these separate
# from the fuzzy third-party aliases above: a localized product style is an
# exact contract, while broad substring matching can misclassify ordinary
# user-created styles.  The real-template regression test verifies this table
# against both ``i18n/locales/*.toml`` and the bundled DOCX templates.
_LOCALIZED_CODE_BLOCK_STYLE_NAMES = frozenset(
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
    )
)
_LOCALIZED_INLINE_CODE_STYLE_NAMES = frozenset(
    (
        "Inline-Code",
        "Inline Code",
        "Código en línea",
        "Code en ligne",
        "インラインコード",
        "인라인 코드",
        "Código em linha",
        "Встроенный код",
        "Mã nội tuyến",
        "行内代码",
        "行內代碼",
    )
)
_LOCALIZED_QUOTE_STYLE_NAMES = {
    f"{stem} {level}": level
    for stem in (
        "Zitat",
        "Quote",
        "Cita",
        "Citation",
        "引用",
        "인용",
        "Citação",
        "Цитата",
        "Trích dẫn",
    )
    for level in range(1, 10)
}
_LEGACY_EXACT_CODE_BLOCK_STYLE_NAMES = frozenset({"code", "source", "programming", "program", "代码", "程序"})
_LEGACY_EXACT_QUOTE_STYLE_NAMES = frozenset({"blockquote"})
_LOCALIZED_CODE_BLOCK_STYLE_NAMES_CASEFOLDED = frozenset(name.casefold() for name in _LOCALIZED_CODE_BLOCK_STYLE_NAMES)
_LOCALIZED_QUOTE_STYLE_NAMES_CASEFOLDED = {
    name.casefold(): level for name, level in _LOCALIZED_QUOTE_STYLE_NAMES.items()
}


@dataclass(frozen=True, slots=True)
class StyleDetectorConfig:
    """Configurable aliases for paragraph style detection (M2).

    Exact aliases are kept separate from optional fuzzy keywords so disabling
    fuzzy matching cannot silently turn aliases back into substring matches.
    """

    code_block_style_names: frozenset[str] = frozenset()
    """Exact paragraph-style aliases that identify a code block."""

    code_block_style_fragments: tuple[str, ...] = ()
    """Extra style-name substrings that identify a code block (case-insensitive)."""

    code_fuzzy_match_enabled: bool = True
    """Enable built-in and configured fuzzy code-style matching."""

    quote_style_names: tuple[tuple[str, int], ...] = ()
    """Exact ``(style_name, level)`` aliases for levelled quote styles."""

    quote_style_patterns: tuple[tuple[str, int], ...] = ()
    """Extra ``(style_name_fragment, level)`` pairs for quote detection."""

    quote_generic_names: frozenset[str] = frozenset()
    """Exact generic quote style aliases (level 1)."""

    quote_fuzzy_match_enabled: bool = True
    """Enable built-in and configured fuzzy quote-style matching."""

    code_character_style_names: frozenset[str] = frozenset()
    """Exact character-style names that identify inline code."""

    quote_character_style_names: frozenset[str] = frozenset()
    """Exact character-style names that identify inline quote text."""

    code_full_paragraph_as_block: bool = True
    """Treat an all-code-run paragraph as a fenced code block."""

    quote_full_paragraph_as_block: bool = True
    """Treat an all-quote-run paragraph as a quote block."""

    wps_shading_enabled: bool = True
    """Recognize WPS solid gray shading as code formatting."""

    word_shading_enabled: bool = True
    """Recognize Word percentage-pattern shading as code formatting."""


@dataclass(frozen=True)
class DocxMarkdownFormattingConfig:
    """Runtime DOCX->MD formatting preservation switches."""

    preserve_formatting: bool = True
    preserve_heading_formatting: bool = False
    preserve_table_header_formatting: bool = False


@dataclass(frozen=True)
class DocxMarkdownSyntaxConfig:
    """Runtime DOCX->MD Markdown syntax choices."""

    bold: str = "asterisk"
    italic: str = "asterisk"
    strikethrough: str = "extended"
    highlight: str = "extended"
    superscript: str = "html"
    subscript: str = "html"
    unordered_list: str = "dash"
    indent_spaces: int = 4


def _build_code_block_fragments(
    config: StyleDetectorConfig | None,
) -> list[str]:
    """Merge built-in code-block fragments with caller-provided aliases."""
    if config is not None and not config.code_fuzzy_match_enabled:
        return []
    if config is None or not config.code_block_style_fragments:
        return list(_CODE_BLOCK_STYLE_FRAGMENTS)
    merged = list(_CODE_BLOCK_STYLE_FRAGMENTS)
    for frag in config.code_block_style_fragments:
        lower = frag.lower().strip()
        if lower and lower not in merged:
            merged.append(lower)
    return merged


def _build_quote_patterns(
    config: StyleDetectorConfig | None,
) -> tuple[list[tuple[re.Pattern, Any]], set[str]]:
    """Merge built-in quote patterns with caller-provided aliases."""
    patterns = list(_QUOTE_STYLE_PATTERNS) if config is None or config.quote_fuzzy_match_enabled else []
    names = set(_QUOTE_GENERIC_NAMES)

    if config is not None:
        if config.quote_fuzzy_match_enabled:
            for raw_frag, level in config.quote_style_patterns:
                frag = raw_frag.strip()
                if frag:
                    escaped = re.escape(frag)
                    if frag.isascii():
                        escaped = rf"(?<![0-9A-Za-z]){escaped}(?:\s*(\d))?(?![0-9A-Za-z])"
                    else:
                        escaped = rf"{escaped}(?:\s*(\d))?"
                    pat = re.compile(escaped, re.IGNORECASE)
                    patterns.append((pat, lambda m, lvl=level: int(m.group(1)) if m.group(1) else lvl))
        for name in config.quote_generic_names:
            lower = name.strip().casefold()
            if lower:
                names.add(lower)

    return patterns, names


def _choice(value: Any, *, allowed: set[str], default: str) -> str:
    if not isinstance(value, str):
        return default
    normalized = value.strip().lower()
    return normalized if normalized in allowed else default


def _int_choice(value: Any, *, allowed: set[int], default: int) -> int:
    try:
        normalized = int(value)
    except (TypeError, ValueError):
        return default
    return normalized if normalized in allowed else default


def docx_markdown_formatting_config_from_conversion_config(
    conversion_config: Mapping[str, Any] | None,
) -> DocxMarkdownFormattingConfig:
    """Build DOCX->MD formatting switches from ``conversion.toml`` data."""
    if not isinstance(conversion_config, Mapping):
        return DocxMarkdownFormattingConfig()
    docx_to_md = conversion_config.get("docx_to_md", {})
    if not isinstance(docx_to_md, Mapping):
        return DocxMarkdownFormattingConfig()
    return DocxMarkdownFormattingConfig(
        preserve_formatting=bool(docx_to_md.get("preserve_formatting", True)),
        preserve_heading_formatting=bool(docx_to_md.get("preserve_heading_formatting", False)),
        preserve_table_header_formatting=bool(docx_to_md.get("preserve_table_header_formatting", False)),
    )


def docx_markdown_syntax_config_from_conversion_config(
    conversion_config: Mapping[str, Any] | None,
) -> DocxMarkdownSyntaxConfig:
    """Build DOCX->MD inline syntax choices from ``conversion.toml`` data."""
    if not isinstance(conversion_config, Mapping):
        return DocxMarkdownSyntaxConfig()
    syntax = conversion_config.get("syntax", {})
    if not isinstance(syntax, Mapping):
        return DocxMarkdownSyntaxConfig()

    marker_choices = {"asterisk", "underscore"}
    extended_choices = {"extended", "html"}
    unordered_choices = {"dash", "asterisk", "plus"}
    return DocxMarkdownSyntaxConfig(
        bold=_choice(syntax.get("bold"), allowed=marker_choices, default="asterisk"),
        italic=_choice(syntax.get("italic"), allowed=marker_choices, default="asterisk"),
        strikethrough=_choice(syntax.get("strikethrough"), allowed=extended_choices, default="extended"),
        highlight=_choice(syntax.get("highlight"), allowed=extended_choices, default="extended"),
        superscript=_choice(syntax.get("superscript"), allowed=extended_choices, default="html"),
        subscript=_choice(syntax.get("subscript"), allowed=extended_choices, default="html"),
        unordered_list=_choice(syntax.get("unordered_list"), allowed=unordered_choices, default="dash"),
        indent_spaces=_int_choice(syntax.get("indent_spaces"), allowed={2, 4}, default=4),
    )


def _string_list(value: Any) -> list[str]:
    if not isinstance(value, list):
        return []
    return [str(item).strip() for item in value if str(item).strip()]


def _casefold_unique_strings(value: Any) -> tuple[str, ...]:
    """Keep the first configured spelling for each case-insensitive alias."""
    by_key: dict[str, str] = {}
    for item in _string_list(value):
        by_key.setdefault(item.casefold(), item)
    return tuple(by_key.values())


def style_detector_config_from_document_config(
    document_config: Mapping[str, Any] | None,
) -> StyleDetectorConfig | None:
    """Build DOCX->MD style detector config from ``document.toml`` data."""
    if not isinstance(document_config, Mapping):
        return None

    style = document_config.get("style", {})
    if not isinstance(style, Mapping):
        return None

    code_docx = style.get("code", {})
    code_docx = code_docx.get("docx_to_md", {}) if isinstance(code_docx, Mapping) else {}
    quote_docx = style.get("quote", {})
    quote_docx = quote_docx.get("docx_to_md", {}) if isinstance(quote_docx, Mapping) else {}

    code_aliases = _casefold_unique_strings(
        code_docx.get("paragraph_style_aliases") if isinstance(code_docx, Mapping) else []
    )
    code_fuzzy_enabled = bool(code_docx.get("fuzzy_match_enabled", True)) if isinstance(code_docx, Mapping) else True
    code_fuzzy_keywords = (
        _casefold_unique_strings(code_docx.get("fuzzy_keywords", []))
        if isinstance(code_docx, Mapping) and code_fuzzy_enabled
        else []
    )

    quote_style_names_by_key: dict[str, tuple[str, int]] = {}
    quote_patterns: list[tuple[str, int]] = []
    quote_generic_names: tuple[str, ...] = ()
    quote_fuzzy_enabled = bool(quote_docx.get("fuzzy_match_enabled", True)) if isinstance(quote_docx, Mapping) else True
    if isinstance(quote_docx, Mapping):
        level_aliases = quote_docx.get("level_style_aliases", {})
        if isinstance(level_aliases, Mapping):
            for raw_name, raw_level in level_aliases.items():
                name = str(raw_name).strip()
                if not name:
                    continue
                try:
                    level = max(1, min(9, int(raw_level)))
                except (TypeError, ValueError):
                    level = 1
                quote_style_names_by_key.setdefault(name.casefold(), (name, level))
        quote_generic_names = _casefold_unique_strings(quote_docx.get("paragraph_style_aliases", []))
        if quote_fuzzy_enabled:
            quote_patterns.extend((name, 1) for name in _casefold_unique_strings(quote_docx.get("fuzzy_keywords", [])))

    code_character_names = _casefold_unique_strings(
        code_docx.get("character_style_aliases", []) if isinstance(code_docx, Mapping) else []
    )
    quote_character_names = _casefold_unique_strings(
        quote_docx.get("character_style_aliases", []) if isinstance(quote_docx, Mapping) else []
    )
    quote_style_names = tuple(quote_style_names_by_key.values())
    has_style_config = isinstance(code_docx, Mapping) or isinstance(quote_docx, Mapping)
    if (
        not has_style_config
        and not code_aliases
        and not code_fuzzy_keywords
        and not quote_style_names
        and not quote_patterns
        and not quote_generic_names
        and not code_character_names
        and not quote_character_names
    ):
        return None
    code_shading = code_docx.get("shading", {}) if isinstance(code_docx, Mapping) else {}
    if not isinstance(code_shading, Mapping):
        code_shading = {}
    return StyleDetectorConfig(
        code_block_style_names=frozenset(code_aliases),
        code_block_style_fragments=tuple(code_fuzzy_keywords),
        code_fuzzy_match_enabled=code_fuzzy_enabled,
        quote_style_names=quote_style_names,
        quote_style_patterns=tuple(quote_patterns),
        quote_generic_names=frozenset(quote_generic_names),
        quote_fuzzy_match_enabled=quote_fuzzy_enabled,
        code_character_style_names=frozenset(code_character_names),
        quote_character_style_names=frozenset(quote_character_names),
        code_full_paragraph_as_block=bool(code_docx.get("full_paragraph_as_block", True)),
        quote_full_paragraph_as_block=bool(quote_docx.get("full_paragraph_as_block", True)),
        wps_shading_enabled=bool(code_shading.get("wps_enabled", True)),
        word_shading_enabled=bool(code_shading.get("word_enabled", True)),
    )


def detect_run_style_type(
    run: Any,
    config: StyleDetectorConfig | None = None,
) -> str | None:
    """Classify one character style as inline ``code`` or ``quote``."""
    try:
        style = getattr(run, "style", None)
        style_name = getattr(style, "name", "") or ""
    except Exception:
        return None
    if not style_name:
        return None

    normalized = style_name.strip().casefold()
    code_names = {name.strip().casefold() for name in _LOCALIZED_INLINE_CODE_STYLE_NAMES}
    quote_names: set[str] = set()
    if config is not None:
        code_names.update(name.strip().casefold() for name in config.code_character_style_names)
        quote_names.update(name.strip().casefold() for name in config.quote_character_style_names)
    if normalized in code_names:
        return "code"
    if normalized in quote_names:
        return "quote"

    probe = type("_StyleProbe", (), {"style": type("_Style", (), {"name": style_name})()})()
    paragraph_type, _level = detect_paragraph_style_type(probe, config=config)
    if paragraph_type == "code_block":
        return "code"
    if paragraph_type == "quote":
        return "quote"
    return None


def detect_paragraph_style_type(
    para: Any,
    config: StyleDetectorConfig | None = None,
) -> tuple[str | None, int | None]:
    """Detect paragraph style as *code_block*, *quote*, or normal.

    Args:
        para: A python-docx Paragraph object.
        config: Optional ``StyleDetectorConfig`` with additional aliases
            that are merged with the built-in defaults.

    Returns:
        ``("code_block", True)`` for code-block styles,
        ``("quote", level)`` for quote styles (level 1–9), or
        ``(None, None)`` for normal paragraphs.
    """
    try:
        style_name = para.style.name if para.style else None
        if not style_name:
            return None, None
    except Exception:
        return None, None

    style_lower = style_name.casefold()
    code_fragments = _build_code_block_fragments(config)
    quote_patterns, quote_names = _build_quote_patterns(config)

    # 1. Exact names created by every bundled locale template.
    if (
        style_lower in _LOCALIZED_CODE_BLOCK_STYLE_NAMES_CASEFOLDED
        or style_lower in _LEGACY_EXACT_CODE_BLOCK_STYLE_NAMES
    ):
        return "code_block", True
    localized_quote_level = _LOCALIZED_QUOTE_STYLE_NAMES_CASEFOLDED.get(style_lower)
    if localized_quote_level is not None:
        return "quote", localized_quote_level
    if style_lower in _LEGACY_EXACT_QUOTE_STYLE_NAMES:
        return "quote", 1

    if config is not None:
        code_names = {name.strip().casefold() for name in config.code_block_style_names}
        if style_lower in code_names:
            return "code_block", True
        quote_levels = {
            name.strip().casefold(): max(1, min(9, level)) for name, level in config.quote_style_names if name.strip()
        }
        configured_quote_level = quote_levels.get(style_lower)
        if configured_quote_level is not None:
            return "quote", configured_quote_level

    # 2. Code block styles
    for fragment in code_fragments:
        normalized_fragment = fragment.strip().casefold()
        if not normalized_fragment:
            continue
        if normalized_fragment.isascii():
            matches = re.search(
                rf"(?<![0-9A-Za-z]){re.escape(normalized_fragment)}(?![0-9A-Za-z])",
                style_lower,
            )
        else:
            matches = normalized_fragment in style_lower
        if matches:
            return "code_block", True

    # 3. Quote styles – try level-specific patterns first
    for pat, level_fn in quote_patterns:
        m = pat.search(style_name)
        if m:
            level = level_fn(m)
            return "quote", max(1, min(9, level))

    # 4. Generic quote names (default level 1)
    if style_lower in {name.casefold() for name in quote_names}:
        return "quote", 1

    return None, None


# ── Gray shading detection (WPS / Word code‑block fallback) ──────────────

# WPS fill colours treated as gray code‑block shading.
_WPS_SHADING_GRAY_COLORS = frozenset(
    {
        "D9D9D9",
        "E7E6E6",
        "F2F2F2",
        "CCCCCC",
        "C0C0C0",
        "A6A6A6",
        "BFBFBF",
        "D0CECE",
    }
)


def _matches_code_shading(
    shading: Any,
    *,
    wps_enabled: bool,
    word_enabled: bool,
) -> bool:
    fill = (shading.get(f"{{{NS_W}}}fill") or "").upper()
    value = (shading.get(f"{{{NS_W}}}val") or "").lower()
    return (wps_enabled and fill in _WPS_SHADING_GRAY_COLORS) or (
        word_enabled and value.startswith("pct") and fill in {"FFFFFF", "AUTO", ""}
    )


def has_gray_shading(
    run: Any,
    *,
    wps_enabled: bool = True,
    word_enabled: bool = True,
) -> bool:
    """Return True if the run has gray character shading (WPS or Word)."""
    try:
        rPr = run._r.find(f"{{{NS_W}}}rPr")
        if rPr is None:
            return False
        shd = rPr.find(f"{{{NS_W}}}shd")
        if shd is None:
            return False
        return _matches_code_shading(
            shd,
            wps_enabled=wps_enabled,
            word_enabled=word_enabled,
        )
    except Exception:
        return False


def has_paragraph_gray_shading(
    para: Any,
    *,
    wps_enabled: bool = True,
    word_enabled: bool = True,
) -> bool:
    """Return True if the paragraph has gray paragraph shading (code block)."""
    try:
        pPr = para._p.find(f"{{{NS_W}}}pPr")
        if pPr is None:
            return False
        shd = pPr.find(f"{{{NS_W}}}shd")
        if shd is None:
            return False
        return _matches_code_shading(
            shd,
            wps_enabled=wps_enabled,
            word_enabled=word_enabled,
        )
    except Exception:
        return False


def detect_full_paragraph_run_style(
    para: Any,
    config: StyleDetectorConfig | None = None,
) -> tuple[str | None, int | bool | None]:
    """Classify paragraphs whose non-empty runs share one character style."""
    try:
        runs = [run for run in para.runs if run.text and run.text.strip()]
    except Exception:
        return None, None
    if not runs:
        return None, None

    effective = config or StyleDetectorConfig()
    classifications: list[str | None] = []
    for run in runs:
        run_type = detect_run_style_type(run, config=effective)
        if run_type is None and has_gray_shading(
            run,
            wps_enabled=effective.wps_shading_enabled,
            word_enabled=effective.word_shading_enabled,
        ):
            run_type = "code"
        classifications.append(run_type)

    if effective.code_full_paragraph_as_block and all(run_type == "code" for run_type in classifications):
        return "code_block", True
    if effective.quote_full_paragraph_as_block and all(run_type == "quote" for run_type in classifications):
        return "quote", 1
    return None, None


# ── Code‑block state accumulator ─────────────────────────────────────────


class CodeBlockAccumulator:
    """Progressive code‑block accumulator for DOCX→MD conversion.

    Collects lines while *in_code_block* is True and emits a fenced
    code block on ``finalize()``.  Indentation is controlled by a
    list‑context level passed to ``start()``.
    """

    def __init__(self, indent_spaces: int = 4) -> None:
        self._indent_spaces = indent_spaces
        self.in_code_block: bool = False
        self._lines: list[str] = []
        self._list_level: int = 0

    def start(self, list_level: int = 0) -> None:
        """Begin accumulating code-block lines."""
        self.in_code_block = True
        self._lines = []
        self._list_level = list_level

    def add_line(self, text: str) -> None:
        """Add a raw line to the code block."""
        if self.in_code_block:
            self._lines.append(text)

    def finalize(self) -> str | None:
        """End the code block and return the fenced Markdown, or None."""
        if not self.in_code_block or not self._lines:
            self.in_code_block = False
            self._lines = []
            return None
        indent = " " * self._indent_spaces * self._list_level
        fence = f"{indent}```"
        physical_lines = [physical_line for line in self._lines for physical_line in re.split(r"\r\n?|\n", line)]
        body = "\n".join(f"{indent}{line}" for line in physical_lines)
        result = f"{fence}\n{body}\n{fence}"
        self.in_code_block = False
        self._lines = []
        return result
