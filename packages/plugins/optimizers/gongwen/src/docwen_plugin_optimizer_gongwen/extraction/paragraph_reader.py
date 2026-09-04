"""Read DOCX document paragraphs into ParagraphFeature objects."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from docx import Document

    from docwen_core.protocols.execution_context import ProgressSink

from docx.oxml.ns import qn

from docwen_core.docx_parsing.image_extraction import extract_images_from_element
from docwen_core.docx_parsing.textbox_extraction import (
    ExtractedParagraph,
    extract_textbox_paragraphs,
)
from docwen_core.text.heading_numbering import detect_heading_prefix
from docwen_plugin_document.shared.list_processing import ListCounterManager
from docwen_plugin_optimizer_gongwen.extraction.format_features import (
    extract_alignment,
    extract_font_info,
    extract_outline_level,
)
from docwen_plugin_optimizer_gongwen.extraction.special_content import (
    detect_table_context,
    has_image,
    has_omml_formula,
    has_page_break,
    has_section_break,
    is_in_textbox,
)
from docwen_plugin_optimizer_gongwen.extraction.table_extraction import (
    extract_table_paragraphs,
)
from docwen_plugin_optimizer_gongwen.models import ParagraphFeature

_GONGWEN_NATIVE_HEADING_FORMATS = frozenset({"chinesecounting", "chinesecountingthousand", "ideographtraditional"})


def _is_native_heading_level(
    level_info: Any,
    *,
    text: str = "",
    style_name: str = "",
    para: Any = None,
    numbering_index: Any = None,
    counter_key: str = "",
    ilvl: int = 0,
    neighboring_paragraphs: tuple[Any, ...] = (),
) -> bool:
    """Accept explicit headings and fail closed for ordinary Chinese lists."""

    style_id = str(getattr(level_info, "p_style", "") or "").casefold().replace(" ", "")
    if any(marker in style_id for marker in ("heading", "title", "标题")):
        return True
    if int(getattr(level_info, "ilvl", -1)) != 0:
        return False
    num_fmt = str(getattr(level_info, "num_fmt", "") or "").casefold()
    lvl_text = str(getattr(level_info, "lvl_text", "") or "").strip().replace(" ", "")
    if num_fmt not in _GONGWEN_NATIVE_HEADING_FORMATS or lvl_text not in {"%1、", "%1．", "%1."}:
        return False

    stripped = text.strip()
    if not stripped or len(stripped) > 40 or stripped.endswith(("。", "；", "！", "？", ";", "!", "?")):
        return False

    normalized_style = style_name.casefold().replace(" ", "")
    if any(marker in normalized_style for marker in ("heading", "title", "标题")):
        return True
    if para is not None:
        outline_level = extract_outline_level(para)
        if isinstance(outline_level, int) and 0 <= outline_level <= 4:
            return True
        if para.paragraph_format.keep_with_next is True:
            return True
        if any(run.bold is True for run in para.runs if run.text.strip()):
            return True

    if numbering_index is None or not neighboring_paragraphs:
        return False
    for neighbor in neighboring_paragraphs:
        try:
            resolved = _resolve_word_numbering(neighbor, numbering_index)
        except Exception:
            return False
        if resolved is not None and resolved[0] == counter_key and resolved[1] == ilvl:
            return False
    return True


def _neighboring_content_paragraphs(paragraphs: list[Any], index: int) -> tuple[Any, ...]:
    """Return the nearest non-empty paragraph on each side."""

    neighbors: list[Any] = []
    for step in (-1, 1):
        candidate_index = index + step
        while 0 <= candidate_index < len(paragraphs):
            candidate = paragraphs[candidate_index]
            if candidate.text.strip():
                neighbors.append(candidate)
                break
            candidate_index += step
    return tuple(neighbors)


def read_paragraphs(
    doc: Document,  # pyright: ignore[reportGeneralTypeIssues]
    output_dir: str | None = None,
    numbering_index=None,
    cleanup_rules: Any = (),
    table_merge_strategy: str = "fill",
    diagnostic_sink: ProgressSink | None = None,
) -> list[ParagraphFeature]:
    """Extract features from all paragraphs in a DOCX document.

    Skips empty/whitespace-only paragraphs.
    Populates rich content fields (formulas, images, breaks, headings).
    Extracts embedded images to the caller-owned *output_dir*.  Text-only
    documents do not require a directory; documents with embedded images fail
    closed when no owned destination was supplied.
    """
    features: list[ParagraphFeature] = []
    numbering_counters = ListCounterManager()

    body_element = doc.element.body
    paragraphs = list(doc.paragraphs)
    for paragraph_index, para in enumerate(paragraphs):
        text = para.text

        # Empty layout paragraphs carry no semantic content, but an empty
        # ``Paragraph.text`` can still own an image, formula or explicit page /
        # section break.  Detect those before deciding whether to discard it.
        has_img = has_image(para)
        has_fml = has_omml_formula(para)
        has_pb = has_page_break(para)
        has_sb = has_section_break(para)
        if (not text or not text.strip()) and not any((has_img, has_fml, has_pb, has_sb)):
            continue

        font_name, font_size_pt = extract_font_info(para)
        style_name = para.style.name if para.style else ""
        outline_level = extract_outline_level(para)
        alignment = extract_alignment(para)

        # ── Special content detection ──
        in_tb = is_in_textbox(para, doc)
        tbl_ctx = detect_table_context(para)

        # ── Formula info ──
        from docwen_plugin_optimizer_gongwen.formula_utils import extract_formula_info

        fml_detected, fml_type, fml_latex = extract_formula_info(para)

        # ── Heading detection ──
        cleaned_text, heading_lvl, heading_num = _detect_heading(
            text,
            style_name,
            para,
            numbering_index=numbering_index,
            numbering_counters=numbering_counters,
            cleanup_rules=cleanup_rules,
            diagnostic_sink=diagnostic_sink,
            diagnostic_location=f"paragraph {len(features)}",
            neighboring_paragraphs=_neighboring_content_paragraphs(paragraphs, paragraph_index),
        )

        # ── Heading clean-text: display text vs raw text ──
        raw_text = text.strip()
        display_text = cleaned_text if heading_lvl > 0 else raw_text
        heading_body_boundary, heading_body_boundary_source = _detect_mixed_heading_body_boundary(
            para,
            display_text,
            raw_text,
            heading_num,
            heading_level=heading_lvl,
        )

        # ── Image extraction ──
        images: list[str] = []
        if has_img:
            if not output_dir:
                raise ValueError("gongwen_image_output_dir_required")
            images = [
                info.path
                for info in extract_images_from_element(
                    para._p,
                    para.part.related_parts,
                    output_dir,
                    name_prefix="gongwen-image",
                )
            ]

        pf = ParagraphFeature(
            index=len(features),
            text=display_text,
            raw_text=raw_text,
            font_name=font_name,
            font_size_pt=font_size_pt,
            style_name=style_name,
            outline_level=outline_level,
            alignment=alignment,
            is_in_textbox=in_tb,
            has_image=has_img,
            has_formula=has_fml or fml_detected,
            formula_type=fml_type,
            formula_latex=fml_latex,
            has_page_break=has_pb,
            has_section_break=has_sb,
            heading_level=heading_lvl,
            heading_numbering_text=heading_num,
            heading_body_boundary=heading_body_boundary,
            heading_body_boundary_source=heading_body_boundary_source,
            extracted_images=images,
            table_cell_context=tbl_ctx,
            source_index=body_element.index(para._p),
        )
        features.append(pf)

    # ── Inject textbox paragraphs after their anchors ────────────────
    textbox_items = extract_textbox_paragraphs(doc)
    if textbox_items:
        tb_features = _textbox_to_features(textbox_items)
        features = _interleave_features(features, tb_features)

    # ── Inject table cell paragraphs after their anchors ─────────────
    table_items = extract_table_paragraphs(
        doc,
        table_merge_strategy=table_merge_strategy,
    )
    if table_items:
        tbl_features = _table_to_features(table_items)
        features = _interleave_features(features, tbl_features)

    # Recognition, re-evaluation and rendering all use ``ParagraphFeature.index``
    # as a list index.  Keep it dense regardless of skipped empty paragraphs
    # and regardless of whether the document contained injectable content.
    _reindex_features(features)
    return features


def _detect_heading(
    text: str,
    style_name: str = "",
    para=None,
    numbering_index=None,
    numbering_counters: ListCounterManager | None = None,
    cleanup_rules: Any = (),
    diagnostic_sink: ProgressSink | None = None,
    diagnostic_location: str = "",
    neighboring_paragraphs: tuple[Any, ...] = (),
) -> tuple[str, int, str]:
    """Three-pass heading detection for body paragraphs.

    Pass 1: Detect heading numbering prefix via shared core rules
            (``docwen_core.text.heading_numbering.detect_heading_prefix``).
            Covers: Chinese numerals, Arabic digits, circled numbers,
            hierarchical numbers (1.1), legal-unit (第x章), etc.
    Pass 2: Check Word pStyle (Heading 1-5).
    Pass 3 (fallback): Try Word numbering definitions via NumberingIndex.

    Returns (cleaned_text, heading_level, numbering_text) where
    heading_level is 1-5 or 0 if not a heading.
    """
    # Pass 1: Shared heading numbering detection
    info = detect_heading_prefix(text, rules=cleanup_rules)
    # A prefix that consumes the whole paragraph (for example a pure numeric
    # 份号 such as ``001``) is metadata, not an empty heading.
    if info is not None and info.clean_text.strip():
        return info.clean_text.strip(), info.numbering_level, info.prefix

    # Pass 2: Word style
    style_map = {"Heading 1": 1, "Heading 2": 2, "Heading 3": 3, "Heading 4": 4, "Heading 5": 5}
    heading_level = style_map.get(style_name, 0)
    heading_num = ""

    # Pass 3: Word numbering definitions via NumberingIndex.  This also runs
    # for style-recognised headings so Word's generated marker is retained.
    if numbering_index is not None and para is not None:
        try:
            resolved = _resolve_word_numbering(para, numbering_index)
            if resolved is not None:
                counter_key, ilvl, level_info = resolved
                counters = numbering_counters or ListCounterManager()
                counter_value = counters.next(counter_key, ilvl, start=level_info.start)
                if _is_native_heading_level(
                    level_info,
                    text=text,
                    style_name=style_name,
                    para=para,
                    numbering_index=numbering_index,
                    counter_key=counter_key,
                    ilvl=ilvl,
                    neighboring_paragraphs=neighboring_paragraphs,
                ):
                    heading_level = heading_level or ilvl + 1
                    heading_num = numbering_index.render_numbering_text(
                        level_info,
                        counter_value,
                        counters.snapshot(counter_key),
                    )
        except Exception as exc:
            if diagnostic_sink is not None:
                diagnostic_sink.report_diagnostic(
                    "warning",
                    (
                        "Could not resolve Word numbering for Gongwen heading detection; "
                        f"kept the paragraph and used style/text fallback after {type(exc).__name__}."
                    ),
                    code="GONGWEN-HEADING-NUMBERING-FALLBACK",
                    location=diagnostic_location,
                )

    return text, heading_level, heading_num


def _resolve_word_numbering(para: Any, numbering_index: Any) -> tuple[str, int, Any] | None:
    """Resolve direct or style-inherited numbering without hiding failures."""

    p_pr = para._p.find(qn("w:pPr"))
    num_pr = p_pr.find(qn("w:numPr")) if p_pr is not None else None
    if num_pr is not None:
        num_id_element = num_pr.find(qn("w:numId"))
        level_element = num_pr.find(qn("w:ilvl"))
        if num_id_element is None:
            return None
        num_id = str(num_id_element.get(qn("w:val"), "0"))
        if num_id == "0":
            return None
        ilvl = int(level_element.get(qn("w:val"), "0")) if level_element is not None else 0
        level_info = numbering_index.lookup(num_id, ilvl)
        return (num_id, ilvl, level_info) if level_info is not None else None

    style_id = getattr(getattr(para, "style", None), "style_id", None)
    if not isinstance(style_id, str) or not style_id:
        return None
    level_info = numbering_index.lookup_by_style_id(style_id)
    if level_info is None:
        return None
    counter_key = f"abs_{level_info.abstract_num_id}"
    return counter_key, int(level_info.ilvl), level_info


_MIXED_HEADING_PUNCTUATION = ("。", "．", ".", "：", ":", "！", "!", "？", "?")


def _detect_mixed_heading_body_boundary(
    para: Any,
    cleaned_text: str,
    raw_text: str,
    heading_numbering_text: str,
    *,
    heading_level: int,
) -> tuple[int | None, str]:
    """Find a Gongwen mixed heading/body boundary after heading recognition.

    Direct run formatting is authoritative when the source exposes it.
    Single-run documents cannot expose a typography boundary, so configured
    Gongwen punctuation provides the explicit plain-text fallback.
    """

    if heading_level <= 0 or not cleaned_text:
        return None, ""

    raw_boundary = _run_format_boundary(para)
    if raw_boundary is not None:
        prefix_length = len(heading_numbering_text) if raw_text.startswith(heading_numbering_text) else 0
        cleaned_boundary = raw_boundary - prefix_length
        if 0 < cleaned_boundary < len(cleaned_text):
            return cleaned_boundary, "run_format"

    boundary = _first_punctuation_heading_boundary(cleaned_text)
    if boundary is not None:
        return boundary, "punctuation_fallback"
    return None, ""


def _run_format_boundary(para: Any) -> int | None:
    """Return the best run boundary that changes from heading to body formatting.

    Inline emphasis can occur inside a heading immediately after sentence
    punctuation.  Such a weak character-style transition must not outrank a
    later direct typography/Normal-body transition that identifies the real
    heading-to-body boundary.
    """

    visible_runs = [run for run in getattr(para, "runs", ()) if run.text]
    if len(visible_runs) < 2:
        return None

    heading_signature = _direct_run_signature(visible_runs[0])
    raw_text = "".join(run.text for run in visible_runs)
    delimiter_boundaries = _heading_delimiter_boundaries(raw_text)
    delimited_candidates: list[tuple[int, bool, int]] = []
    boundary = len(visible_runs[0].text)
    for run_index, run in enumerate(visible_runs[1:], start=1):
        signature = _direct_run_signature(run)
        if _looks_like_body_format_transition(heading_signature, signature):
            if delimiter_boundaries:
                # A title-internal emphasis/code run is not the body.  With a
                # visible heading terminator, the authoritative run boundary
                # must immediately follow one of those terminators (allowing
                # only whitespace in between).  Later body-internal styling is
                # left to the punctuation fallback instead of splitting late.
                if any(
                    delimiter_boundary <= boundary and not raw_text[delimiter_boundary:boundary].strip()
                    for delimiter_boundary in delimiter_boundaries
                ):
                    delimited_candidates.append(
                        (
                            boundary,
                            _is_strong_body_format_transition(
                                heading_signature,
                                signature,
                            ),
                            run_index,
                        )
                    )
            elif _is_strong_body_format_transition(heading_signature, signature):
                return boundary
        boundary += len(run.text)

    strong_candidates = [candidate for candidate in delimited_candidates if candidate[1]]
    if strong_candidates:
        return strong_candidates[0][0]

    # A one-run inline style that immediately returns to the heading format is
    # a title-internal emphasis/code span when a later boundary exists.  Keep a
    # lone weak candidate for legacy documents whose body itself starts with an
    # Emphasis/Inline Code run.
    for position, (candidate_boundary, _strong, run_index) in enumerate(delimited_candidates):
        has_later_candidate = position + 1 < len(delimited_candidates)
        next_returns_to_heading = run_index + 1 < len(visible_runs) and not _looks_like_body_format_transition(
            heading_signature,
            _direct_run_signature(visible_runs[run_index + 1]),
        )
        if has_later_candidate and next_returns_to_heading:
            continue
        return candidate_boundary
    return None


def _heading_delimiter_boundaries(text: str) -> list[int]:
    """Return all non-terminal mixed-heading delimiter boundaries in order."""

    return sorted(
        {
            position + len(delimiter)
            for delimiter in _MIXED_HEADING_PUNCTUATION
            for position in range(len(text))
            if text.startswith(delimiter, position) and position + len(delimiter) < len(text)
        }
    )


def _first_punctuation_heading_boundary(text: str) -> int | None:
    """Choose the earliest valid delimiter by text position, not tuple order."""

    boundaries = _heading_delimiter_boundaries(text)
    return boundaries[0] if boundaries else None


def _direct_run_signature(run: Any) -> dict[str, Any]:
    r_pr = run._r.find(qn("w:rPr"))
    fonts = r_pr.find(qn("w:rFonts")) if r_pr is not None else None
    style_name = run.style.name if run.style is not None else ""
    return {
        "name": run.font.name,
        "east_asia": fonts.get(qn("w:eastAsia")) if fonts is not None else None,
        "size": run.font.size.pt if run.font.size is not None else None,
        "bold": run.bold,
        "italic": run.italic,
        "underline": run.underline,
        "style_name": style_name,
    }


def _looks_like_body_format_transition(heading: dict[str, Any], candidate: dict[str, Any]) -> bool:
    if candidate["bold"] is False and heading["bold"] is not False:
        return True
    for key in ("name", "east_asia", "size"):
        if candidate[key] is not None and candidate[key] != heading[key]:
            return True
    candidate_style = str(candidate["style_name"]).casefold()
    heading_style = str(heading["style_name"]).casefold()
    return candidate_style != heading_style and any(
        token in candidate_style
        for token in (
            "normal",
            "body",
            "正文",
            "emphasis",
            "强调",
            "inline code",
            "code",
            "代码",
        )
    )


def _is_strong_body_format_transition(heading: dict[str, Any], candidate: dict[str, Any]) -> bool:
    """Reject ambiguous inline-only style changes when no terminator exists."""

    if candidate["bold"] is False and heading["bold"] is not False:
        return True
    if any(candidate[key] is not None and candidate[key] != heading[key] for key in ("name", "east_asia", "size")):
        return True
    candidate_style = str(candidate["style_name"]).casefold()
    heading_style = str(heading["style_name"]).casefold()
    return candidate_style != heading_style and any(token in candidate_style for token in ("normal", "body", "正文"))


def _textbox_to_features(
    items: list[ExtractedParagraph],
) -> list[tuple[int | None, ParagraphFeature]]:
    """Convert ExtractedParagraph objects to (anchor_index, ParagraphFeature) pairs."""
    result: list[tuple[int | None, ParagraphFeature]] = []
    for item in items:
        pf = ParagraphFeature(
            index=-1,  # will be reindexed
            text=item.text,
            raw_text=item.text,
            is_in_textbox=True,
            source=item.source,
            source_index=item.anchor_index,
        )
        result.append((item.anchor_index, pf))
    return result


def _table_to_features(
    items,
) -> list[tuple[int | None, ParagraphFeature]]:
    """Convert ExtractedTableParagraph objects to (anchor_index, ParagraphFeature) pairs."""
    result: list[tuple[int | None, ParagraphFeature]] = []
    for item in items:
        pf = ParagraphFeature(
            index=-1,
            text=item.text,
            raw_text=item.text,
            source=item.source,
            table_cell_context=item.table_cell_context,
            source_index=item.anchor_index,
            table_index=item.table_index,
            table_row_index=item.row_index,
            table_cell_index=item.cell_index,
            table_markdown=item.table_markdown,
            is_table_anchor=item.is_table_anchor,
            table_fidelity_risks=item.table_fidelity_risks,
        )
        result.append((item.anchor_index, pf))
    return result


def _interleave_features(
    body: list[ParagraphFeature],
    injected: list[tuple[int | None, ParagraphFeature]],
) -> list[ParagraphFeature]:
    """Insert injected features after their anchor paragraph index.

    Features with anchor_index=None are placed at the start.
    """
    entries: list[tuple[int, int, int, ParagraphFeature]] = []
    sequence = 0
    for pf in body:
        source_index = pf.source_index if pf.source_index is not None else -1
        # A normal body paragraph owns its top-level element and therefore
        # precedes textbox content anchored to that same paragraph.
        same_anchor_order = 0 if pf.source == "body" else 1
        entries.append((source_index, same_anchor_order, sequence, pf))
        sequence += 1

    for anchor, pf in injected:
        pf.source_index = anchor
        entries.append((anchor if anchor is not None else -1, 1, sequence, pf))
        sequence += 1

    entries.sort(key=lambda entry: entry[:3])
    result = [entry[3] for entry in entries]
    _reindex_features(result)
    return result


def _reindex_features(features: list[ParagraphFeature]) -> None:
    """Make the public feature index exactly match its list position."""

    for index, feature in enumerate(features):
        feature.index = index
