"""Run segmentation and Markdown rendering for DOCX paragraphs.

Resolves hyperlink relationship targets, applies inline formatting,
and delegates note references to NoteExtractor.
"""

from __future__ import annotations

from collections.abc import Callable
from typing import Any

from docwen_core.docx_parsing.format_features import (
    DocxMarkdownSyntaxConfig,
    StyleDetectorConfig,
    detect_run_style_type,
)
from docwen_core.docx_parsing.xml_ns import NS_W

# Gray fill colours treated as inline code shading.
_WPS_RUN_SHADING_GRAY_COLORS = frozenset(
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


def _apply_wrappers(text: str, wrappers: list[tuple[str, str]]) -> str:
    for prefix, suffix in wrappers:
        text = f"{prefix}{text}{suffix}"
    return text


def _unwrap_wrappers(text: str, wrappers: list[tuple[str, str]]) -> str | None:
    inner = text
    for prefix, suffix in reversed(wrappers):
        if not inner.startswith(prefix) or not inner.endswith(suffix):
            return None
        inner = inner[len(prefix) : len(inner) - len(suffix)]
    return inner


def _append_wrapped(parts: list[str], raw_text: str, wrappers: list[tuple[str, str]]) -> None:
    rendered = _apply_wrappers(raw_text, wrappers)
    if wrappers and parts:
        previous_inner = _unwrap_wrappers(parts[-1], wrappers)
        if previous_inner is not None:
            parts[-1] = _apply_wrappers(previous_inner + raw_text, wrappers)
            return
    parts.append(rendered)


def _marker_pair(kind: str, config: DocxMarkdownSyntaxConfig) -> tuple[str, str]:
    if kind == "bold":
        marker = "__" if config.bold == "underscore" else "**"
    else:
        marker = "_" if config.italic == "underscore" else "*"
    return marker, marker


def _extended_pair(kind: str, config: DocxMarkdownSyntaxConfig) -> tuple[str, str]:
    if kind == "strikethrough":
        return ("<del>", "</del>") if config.strikethrough == "html" else ("~~", "~~")
    if kind == "highlight":
        return ("<mark>", "</mark>") if config.highlight == "html" else ("==", "==")
    if kind == "superscript":
        return ("^", "^") if config.superscript == "extended" else ("<sup>", "</sup>")
    return ("~", "~") if config.subscript == "extended" else ("<sub>", "</sub>")


def _on_off_property_is_enabled(element: Any) -> bool:
    value = element.get(f"{{{NS_W}}}val")
    return value is None or value.casefold() not in {"0", "false", "off", "no", "none"}


def _format_wrappers(
    *,
    has_shading: bool,
    has_highlight: bool,
    is_superscript: bool,
    is_subscript: bool,
    is_strikethrough: bool,
    is_underline: bool,
    is_bold: bool,
    is_italic: bool,
    syntax_config: DocxMarkdownSyntaxConfig,
) -> list[tuple[str, str]]:
    wrappers: list[tuple[str, str]] = []
    if has_shading:
        wrappers.append(("`", "`"))
    if is_superscript:
        wrappers.append(_extended_pair("superscript", syntax_config))
    if is_subscript:
        wrappers.append(_extended_pair("subscript", syntax_config))
    if is_strikethrough:
        wrappers.append(_extended_pair("strikethrough", syntax_config))
    if is_underline:
        wrappers.append(("<u>", "</u>"))
    if is_bold and is_italic and syntax_config.bold == syntax_config.italic:
        marker = "___" if syntax_config.bold == "underscore" else "***"
        wrappers.append((marker, marker))
    else:
        if is_italic:
            wrappers.append(_marker_pair("italic", syntax_config))
        if is_bold:
            wrappers.append(_marker_pair("bold", syntax_config))
    if has_highlight:
        wrappers.append(_extended_pair("highlight", syntax_config))
    return wrappers


def append_formatted_run_text(
    parts: list[str],
    raw_text: str,
    run: Any,
    *,
    syntax_config: DocxMarkdownSyntaxConfig,
    style_detector_config: StyleDetectorConfig | None = None,
    run_style_type: str | None = None,
) -> None:
    """Append one OOXML run's text using an explicit Markdown syntax policy.

    Formatting is read from the run's ``w:rPr`` element.  Adjacent runs with
    identical wrappers are coalesced so callers do not emit constructs such as
    ``**Hello ****World**``.  ``syntax_config`` is deliberately required: a
    request-scoped caller must not fall back to mutable process-wide policy.

    Args:
        parts: Output fragments accumulated by the paragraph renderer.
        raw_text: Plain text extracted from the run.
        run: The OOXML ``w:r`` element containing optional run properties.
        syntax_config: Markdown markers selected for the current request.
    """
    r_pr = run.find(f"{{{NS_W}}}rPr")
    is_bold = False
    is_italic = False
    is_underline = False
    is_strikethrough = False
    is_superscript = False
    is_subscript = False
    has_highlight = False
    has_shading = False
    if r_pr is not None:
        bold = r_pr.find(f"{{{NS_W}}}b")
        italic = r_pr.find(f"{{{NS_W}}}i")
        underline = r_pr.find(f"{{{NS_W}}}u")
        strike = r_pr.find(f"{{{NS_W}}}strike")
        is_bold = bold is not None and _on_off_property_is_enabled(bold)
        is_italic = italic is not None and _on_off_property_is_enabled(italic)
        is_underline = underline is not None and _on_off_property_is_enabled(underline)
        is_strikethrough = strike is not None and _on_off_property_is_enabled(strike)
        vert = r_pr.find(f"{{{NS_W}}}vertAlign")
        if vert is not None:
            val = vert.get(f"{{{NS_W}}}val")
            if val == "superscript":
                is_superscript = True
            elif val == "subscript":
                is_subscript = True

        highlight = r_pr.find(f"{{{NS_W}}}highlight")
        if highlight is not None:
            highlight_value = highlight.get(f"{{{NS_W}}}val")
            if highlight_value and highlight_value.lower() not in ("none", ""):
                has_highlight = True

        shading = r_pr.find(f"{{{NS_W}}}shd")
        if shading is not None:
            fill = (shading.get(f"{{{NS_W}}}fill") or "").upper()
            value = (shading.get(f"{{{NS_W}}}val") or "").lower()
            effective = style_detector_config or StyleDetectorConfig()
            has_shading = (effective.wps_shading_enabled and fill in _WPS_RUN_SHADING_GRAY_COLORS) or (
                effective.word_shading_enabled and value.startswith("pct") and fill in {"FFFFFF", "AUTO", ""}
            )

    # Quote character styles can promote a whole paragraph to a blockquote,
    # but Markdown has no inline quote marker.  Mapping them to shading would
    # corrupt the text by emitting inline-code ticks.
    has_shading = has_shading or run_style_type == "code"

    wrappers = _format_wrappers(
        has_shading=has_shading,
        has_highlight=has_highlight,
        is_superscript=is_superscript,
        is_subscript=is_subscript,
        is_strikethrough=is_strikethrough,
        is_underline=is_underline,
        is_bold=is_bold,
        is_italic=is_italic,
        syntax_config=syntax_config,
    )
    _append_wrapped(parts, raw_text, wrappers)


def resolve_hyperlink_target(para: Any, hyperlink_element: Any) -> str | None:
    """Resolve the real URL from a ``<w:hyperlink>`` element.

    Uses ``para.part.rels[rel_id].target_ref`` to resolve the
    relationship id stored in ``r:id``.

    Args:
        para: A python-docx Paragraph object.
        hyperlink_element: The ``<w:hyperlink>`` lxml element.

    Returns:
        The resolved URL string, or None if resolution fails.
    """
    r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    rel_id = hyperlink_element.get(f"{{{r_ns}}}id")
    if not rel_id:
        return None

    try:
        part = para.part
        rel = part.rels[rel_id]
        return rel.target_ref
    except (AttributeError, KeyError):
        return None


def _render_paragraph_run_segments(
    para: Any,
    note_extractor: Any = None,
    preserve_formatting: bool = True,
    *,
    syntax_config: DocxMarkdownSyntaxConfig,
    style_detector_config: StyleDetectorConfig | None = None,
    split_page_breaks: bool,
    math_renderer: Callable[[Any], str | None] | None = None,
    root_element: Any | None = None,
) -> list[str]:
    """Render paragraph XML children into page-aware Markdown segments."""
    segments: list[list[str]] = [[]]

    def _parts() -> list[str]:
        return segments[-1]

    def _append(text: str) -> None:
        if text:
            _parts().append(text)

    def _handle_break(child: Any) -> None:
        break_type = child.get(f"{{{NS_W}}}type")
        if split_page_breaks and break_type == "page":
            segments.append([])
        else:
            _parts().append("\n")

    def _process_hyperlink(child: Any) -> None:
        url = resolve_hyperlink_target(para, child) or ""
        hyperlink_segments = _render_paragraph_run_segments(
            para,
            note_extractor=note_extractor,
            preserve_formatting=preserve_formatting,
            syntax_config=syntax_config,
            style_detector_config=style_detector_config,
            split_page_breaks=split_page_breaks,
            math_renderer=math_renderer,
            root_element=child,
        )
        for index, text in enumerate(hyperlink_segments):
            if text:
                _append(f"[{text}]({url})" if url else text)
            if index < len(hyperlink_segments) - 1:
                if split_page_breaks:
                    segments.append([])
                else:
                    _parts().append("\n")

    def _process_children(parent: Any) -> None:
        for child in parent:
            tag = child.tag.split("}")[-1] if child.tag and "}" in child.tag else (child.tag or "")

            if tag == "r":
                _handle_run(child)
            elif tag == "hyperlink":
                _process_hyperlink(child)
            elif tag in ("ins", "moveTo", "fldSimple", "smartTag", "sdt", "sdtContent", "customXml"):
                _process_children(child)
            elif tag in ("del", "moveFrom"):
                pass  # skip tracked deletions
            elif tag in ("oMath", "oMathPara"):
                if math_renderer is not None:
                    _append(math_renderer(child) or "")
                else:
                    # Pass through raw text; formula rendering is external.
                    for text_element in child.iter(f"{{{NS_W}}}t"):
                        _append(text_element.text or "")
            elif tag == "AlternateContent":
                # A modern Choice and compatibility Fallback normally encode
                # the same visible content.  Select one effective branch so we
                # neither drop the wrapper nor duplicate its payload.  If a
                # Choice has no content understood by this renderer, try the
                # next Choice and finally Fallback.
                branches: dict[str, list[Any]] = {"Choice": [], "Fallback": []}
                for branch in child:
                    branch_tag = branch.tag.split("}")[-1] if branch.tag and "}" in branch.tag else (branch.tag or "")
                    if branch_tag in branches:
                        branches[branch_tag].append(branch)
                for branch_kind in ("Choice", "Fallback"):
                    branch_rendered = False
                    for branch in branches[branch_kind]:
                        before = [list(segment) for segment in segments]
                        _process_children(branch)
                        if segments != before:
                            branch_rendered = True
                            break
                    if branch_rendered:
                        break

    def _append_run_text(run: Any, text: str) -> None:
        w_ns = NS_W

        # ── M1: Task list checkbox detection ────────────────────────────
        # Unicode checkbox characters (detect before any formatting)
        checkbox_map_unicode = {
            "\u2610": "- [ ]",
            "\u2611": "- [x]",
            "\u2612": "- [~]",
        }
        if text in checkbox_map_unicode:
            _append(checkbox_map_unicode[text])
            return

        # Wingdings font checkbox characters
        run_properties = run.find(f"{{{w_ns}}}rPr")
        if run_properties is not None:
            run_fonts = run_properties.find(f"{{{w_ns}}}rFonts")
            if run_fonts is not None:
                font_name = (
                    run_fonts.get(f"{{{w_ns}}}ascii")
                    or run_fonts.get(f"{{{w_ns}}}hAnsi")
                    or run_fonts.get(f"{{{w_ns}}}cs")
                    or ""
                )
                if font_name and "wingdings" in font_name.lower():
                    wingdings_map = {
                        "\u00a8": "- [ ]",  # ¨ → unchecked
                        "\u00b0": "- [x]",  # ° → checked
                        "\u00d7": "- [~]",  # × → uncertain
                    }
                    if text in wingdings_map:
                        _append(wingdings_map[text])
                        return

        if not preserve_formatting:
            _append(text)
            return

        append_formatted_run_text(
            _parts(),
            text,
            run,
            syntax_config=syntax_config,
            style_detector_config=style_detector_config,
            run_style_type=_resolved_run_style_type(run),
        )

    def _resolved_run_style_type(run: Any) -> str | None:
        try:
            from docx.text.run import Run

            if not hasattr(para, "part"):
                return None
            return detect_run_style_type(Run(run, para), config=style_detector_config)
        except Exception:
            return None

    def _handle_run(run: Any) -> None:
        w_ns = NS_W
        run_properties = run.find(f"{{{w_ns}}}rPr")
        vanish = run_properties.find(f"{{{w_ns}}}vanish") if run_properties is not None else None
        if vanish is not None and _on_off_property_is_enabled(vanish):
            return
        text_parts: list[str] = []

        def _flush_text() -> None:
            if text_parts:
                _append_run_text(run, "".join(text_parts))
                text_parts.clear()

        for child in run:
            tag = child.tag.split("}")[-1] if child.tag and "}" in child.tag else (child.tag or "")
            if tag == "t" and child.text is not None:
                text_parts.append(child.text)
            elif tag == "tab":
                _flush_text()
                _parts().append("\t")
            elif tag == "noBreakHyphen":
                text_parts.append("\u2011")
            elif tag == "softHyphen":
                text_parts.append("\u00ad")
            elif tag in ("br", "cr"):
                _flush_text()
                _handle_break(child)
            elif tag == "footnoteReference" and note_extractor is not None:
                _flush_text()
                w_id = child.get(f"{{{w_ns}}}id")
                if w_id is not None:
                    _append(note_extractor.get_reference_text("footnote", int(w_id)))
            elif tag == "endnoteReference" and note_extractor is not None:
                _flush_text()
                w_id = child.get(f"{{{w_ns}}}id")
                if w_id is not None:
                    _append(note_extractor.get_reference_text("endnote", int(w_id)))

        _flush_text()

    element = root_element if root_element is not None else getattr(para, "_p", para)
    _process_children(element)
    return ["".join(parts) for parts in segments]


def render_paragraph_runs(
    para: Any,
    note_extractor: Any = None,
    preserve_formatting: bool = True,
    *,
    syntax_config: DocxMarkdownSyntaxConfig,
    style_detector_config: StyleDetectorConfig | None = None,
) -> str:
    """Render paragraph XML children in order, producing Markdown text.

    Handles ``w:hyperlink`` (resolving ``r:id`` to target URL),
    ``w:r`` (with bold/italic/underline/strikethrough formatting),
    and delegating ``w:footnoteReference``/``w:endnoteReference`` to
    ``NoteExtractor.get_reference_text()``.

    Args:
        para: A python-docx Paragraph object.
        note_extractor: Optional NoteExtractor for inline note references.
        preserve_formatting: If True, emit ``**bold**`` etc.
        syntax_config: Inline Markdown syntax policy owned by this request.

    Returns:
        Markdown string for the paragraph content.
    """
    return _render_paragraph_run_segments(
        para,
        note_extractor=note_extractor,
        preserve_formatting=preserve_formatting,
        syntax_config=syntax_config,
        style_detector_config=style_detector_config,
        split_page_breaks=False,
    )[0]


def render_paragraph_runs_split_on_page_breaks(
    para: Any,
    note_extractor: Any = None,
    preserve_formatting: bool = True,
    *,
    syntax_config: DocxMarkdownSyntaxConfig,
    style_detector_config: StyleDetectorConfig | None = None,
    math_renderer: Callable[[Any], str | None] | None = None,
) -> list[str]:
    """Render rich Markdown segments separated at accepted page breaks.

    Unlike the legacy plain-text splitter, this walks the paragraph OOXML in
    document order and therefore retains formatting, hyperlinks, accepted
    revisions, note references, formulas supplied by ``math_renderer``, and
    text on either side of a page break inside the same run. Empty segments are
    retained so the caller can collapse adjacent separators deterministically.
    """
    return _render_paragraph_run_segments(
        para,
        note_extractor=note_extractor,
        preserve_formatting=preserve_formatting,
        syntax_config=syntax_config,
        style_detector_config=style_detector_config,
        split_page_breaks=True,
        math_renderer=math_renderer,
    )
