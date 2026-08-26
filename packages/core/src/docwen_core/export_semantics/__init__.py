"""Immutable export policies and pure rendering helpers for DocWen.

This core module must not import from any plugin or runtime package.
"""

from __future__ import annotations

import logging
from collections.abc import Mapping
from dataclasses import dataclass
from typing import Any

logger = logging.getLogger(__name__)

# ═══════════════════════════════════════════════════════════════════════════
# Constants
# ═══════════════════════════════════════════════════════════════════════════

VALID_TABLE_MERGE_EXPORT_STRATEGIES: tuple[str, ...] = ("fill", "empty", "marker")

VALID_LINK_STYLES: tuple[str, ...] = (
    "wiki_embed",
    "wiki_link",
    "markdown_embed",
    "markdown_link",
)

_DEFAULT_YAML_LIST_SEPARATOR: str = "、"
_DEFAULT_IMAGE_EXTRACTION_MODE: str = "file"
_DEFAULT_OCR_PLACEMENT_MODE: str = "main_md"

# ═══════════════════════════════════════════════════════════════════════════
# Dataclasses
# ═══════════════════════════════════════════════════════════════════════════


@dataclass(frozen=True, slots=True)
class TableMergeRegion:
    """A rectangular merge region in a table.

    F-H3b-027
    """

    start_row: int
    start_col: int
    end_row: int
    end_col: int

    @property
    def rowspan(self) -> int:
        """Number of rows this region spans."""
        return self.end_row - self.start_row + 1

    @property
    def colspan(self) -> int:
        """Number of columns this region spans."""
        return self.end_col - self.start_col + 1


@dataclass(frozen=True, slots=True)
class TableSemanticCell:
    """Semantic cell representation in a merge-aware table grid.

    Each cell carries its position, text content, and merge-region metadata
    so that downstream renderers can apply strategy-specific rendering
    (fill / marker / empty) without duplicating grid-building logic.

    F-H3b-027, F-H3b-029
    """

    row: int
    col: int
    raw_text: str
    display_text: str
    anchor_text: str
    anchor_row: int
    anchor_col: int
    rowspan: int
    colspan: int
    is_anchor: bool
    is_covered: bool

    @property
    def marker(self) -> str | None:
        """Return a visual marker for covered cells.

        ``"<"`` when the covered cell is on the same row as its anchor,
        ``"^"`` when on a different row.  Returns ``None`` for non-covered cells.
        """
        if not self.is_covered:
            return None
        return "<" if self.row == self.anchor_row else "^"


@dataclass(frozen=True, slots=True)
class LinkRuntimeConfig:
    """Immutable link-processing policy for one conversion request."""

    max_depth: int = 3
    non_embed_wiki_mode: str = "hyperlink"
    non_embed_markdown_mode: str = "hyperlink"
    embed_wiki_image_mode: str = "embed"
    embed_markdown_image_mode: str = "embed"
    embed_md_file_mode: str = "embed"
    search_dirs: tuple[str, ...] = (
        ".",
        "assets",
        "images",
        "attachments",
    )
    detect_circular: bool = True
    file_not_found_mode: str = "placeholder"
    circular_reference_mode: str = "placeholder"
    max_depth_reached_mode: str = "placeholder"
    auto_link_bare_url: bool = False

    @classmethod
    def from_config(cls, config: dict[str, Any]) -> LinkRuntimeConfig:
        """Build a ``LinkRuntimeConfig`` from a flat config dictionary.

        The dictionary keys map to the ``link.*`` TOML section keys, e.g.
        ``link.format.image_link_style``, ``link.embedding.max_depth``, etc.
        """
        non_embed = config.get("non_embed_links", {})
        embed_links = config.get("embed_links", {})
        embedding = config.get("embedding", {})
        path_res = config.get("path_resolution", {})
        err_handling = config.get("error_handling", {})

        return cls(
            max_depth=int(embedding.get("max_depth", 3)),
            non_embed_wiki_mode=str(non_embed.get("wiki_mode", "hyperlink")),
            non_embed_markdown_mode=str(non_embed.get("markdown_mode", "hyperlink")),
            auto_link_bare_url=non_embed.get("auto_link_bare_url") is True,
            embed_wiki_image_mode=str(embed_links.get("wiki_image_mode", "embed")),
            embed_markdown_image_mode=str(embed_links.get("markdown_image_mode", "embed")),
            embed_md_file_mode=str(embed_links.get("md_file_mode", "embed")),
            search_dirs=tuple(str(d) for d in path_res.get("search_dirs", (".", "assets", "images", "attachments"))),
            detect_circular=bool(err_handling.get("detect_circular", True)),
            file_not_found_mode=str(err_handling.get("file_not_found", "placeholder")),
            circular_reference_mode=str(err_handling.get("circular_reference", "placeholder")),
            max_depth_reached_mode=str(err_handling.get("max_depth_reached", "placeholder")),
        )


@dataclass(frozen=True, slots=True)
class MarkdownExportSemantics:
    """User-facing markdown export semantic settings.

    Replaces the old dict-based ``_DEFAULT_MARKDOWN_SEMANTICS`` and
    ``_DEFAULT_MARKDOWN_MODES`` with typed, immutable fields.

    F-H3b-001, F-H3b-030, F-H3b-031, F-H3b-033, F-H3b-034
    """

    # Link style (F-H3b-030)
    image_link_style: str = "wiki_embed"
    md_file_link_style: str = "wiki_embed"

    # OCR blockquote title (F-H3b-031)
    ocr_blockquote_title_enabled: bool = True
    ocr_blockquote_title_override_text: str = ""

    # Base64 export (F-H3b-001)
    export_base64_compress_enabled: bool = True
    export_base64_compress_threshold_kb: int = 100

    # Export modes
    image_extraction_mode: str = _DEFAULT_IMAGE_EXTRACTION_MODE
    ocr_placement_mode: str = _DEFAULT_OCR_PLACEMENT_MODE
    table_merge_export_strategy: str = "fill"

    # DOCX -> Markdown break separators
    page_break_separator: str = "---"
    section_break_separator: str = "***"
    horizontal_rule_separator: str = "___"

    # YAML list separator (F-H3b-033)
    yaml_list_separator: str = "、"

    # Intermediate files (F-H3b-034)
    save_intermediate_files: bool = False

    @classmethod
    def from_config_snapshot(
        cls,
        config_snapshot: Mapping[str, Any] | None,
        *,
        requested_locale: object = None,
    ) -> MarkdownExportSemantics:
        """Purely project one immutable request snapshot into export policy."""
        snapshot = config_snapshot if isinstance(config_snapshot, Mapping) else {}
        conversion = _mapping_section(snapshot.get("conversion"))
        link = _mapping_section(snapshot.get("link"))
        output = _mapping_section(snapshot.get("output"))
        locale = _snapshot_locale(snapshot, requested_locale)
        return cls.from_config(
            link_format=_mapping_section(link.get("format")),
            ocr_output=_mapping_section(conversion.get("ocr_output")),
            export_cfg=_mapping_section(snapshot.get("export")),
            conversion_cfg=conversion,
            intermediate_files_cfg=_mapping_section(output.get("intermediate_files")),
            locale=locale,
        )

    @classmethod
    def from_config(
        cls,
        link_format: dict[str, Any] | None = None,
        ocr_output: dict[str, Any] | None = None,
        export_cfg: dict[str, Any] | None = None,
        conversion_cfg: dict[str, Any] | None = None,
        intermediate_files_cfg: dict[str, Any] | None = None,
        locale: str = "zh_CN",
    ) -> MarkdownExportSemantics:
        """Build a ``MarkdownExportSemantics`` from TOML config sections.

        Parameters match the keys under the runtime ``DEFAULT_*_TOML`` dicts.
        """
        link = link_format or {}
        ocr = ocr_output or {}
        export_ = export_cfg or {}
        conv = conversion_cfg or {}
        intermediate = intermediate_files_cfg or {}
        conversion_export = conv.get("export", {})
        if not isinstance(conversion_export, dict):
            conversion_export = {}
        md_to_docx = conv.get("md_to_docx", {})
        if not isinstance(md_to_docx, dict):
            md_to_docx = {}
        yaml_list_separator = md_to_docx.get("list_separator", "、")
        if yaml_list_separator is None:
            yaml_list_separator = "、"

        table_strategy = normalize_table_merge_export_strategy(
            conv.get("table_merge_export_strategy"),
        )
        horizontal_rule_cfg = conv.get("horizontal_rule", {})
        docx_to_md_breaks = horizontal_rule_cfg.get("docx_to_md", {}) if isinstance(horizontal_rule_cfg, dict) else {}
        image_extraction_mode = export_.get("to_md_image_extraction_mode", _DEFAULT_IMAGE_EXTRACTION_MODE)
        ocr_placement_mode = export_.get("to_md_ocr_placement_mode", _DEFAULT_OCR_PLACEMENT_MODE)

        return cls(
            image_link_style=_normalize_link_style(link.get("image_link_style", "wiki_embed")),
            md_file_link_style=_normalize_link_style(link.get("md_file_link_style", "wiki_embed")),
            ocr_blockquote_title_enabled=bool(ocr.get("show_blockquote_title", True)),
            ocr_blockquote_title_override_text=str(
                ocr.get("blockquote_title_override_by_locale", {}).get(locale, "")
                if isinstance(ocr.get("blockquote_title_override_by_locale"), dict)
                else ""
            ),
            export_base64_compress_enabled=bool(conversion_export.get("base64_compress_enabled", True)),
            export_base64_compress_threshold_kb=int(conversion_export.get("base64_compress_threshold_kb", 100)),
            image_extraction_mode=str(image_extraction_mode),
            ocr_placement_mode=str(ocr_placement_mode),
            table_merge_export_strategy=table_strategy,
            page_break_separator=normalize_markdown_break_separator(
                docx_to_md_breaks.get("page_break"),
                default="---",
            ),
            section_break_separator=normalize_markdown_break_separator(
                docx_to_md_breaks.get("section_break"),
                default="***",
            ),
            horizontal_rule_separator=normalize_markdown_break_separator(
                docx_to_md_breaks.get("horizontal_rule"),
                default="___",
            ),
            yaml_list_separator=str(yaml_list_separator),
            save_intermediate_files=bool(intermediate.get("save_to_output", False)),
        )


def _mapping_section(value: object) -> dict[str, Any]:
    if isinstance(value, Mapping):
        return {str(key): item for key, item in value.items()}
    return {}


def _snapshot_locale(snapshot: Mapping[str, Any], requested_locale: object) -> str:
    if isinstance(requested_locale, str) and requested_locale.strip():
        return requested_locale.strip()
    gui = _mapping_section(snapshot.get("gui"))
    language = gui.get("language")
    locale = language.get("locale") if isinstance(language, Mapping) else language
    return str(locale or "zh_CN")


def _title_from_semantics(semantics: MarkdownExportSemantics) -> str:
    if not semantics.ocr_blockquote_title_enabled:
        return ""
    return semantics.ocr_blockquote_title_override_text.strip()


@dataclass(frozen=True, slots=True)
class MarkdownRequestPolicy:
    """Generic Markdown policy frozen for one admitted conversion request."""

    export: MarkdownExportSemantics
    ocr_blockquote_title: str


def resolve_markdown_request_policy(context: object) -> MarkdownRequestPolicy:
    """Resolve one context-owned policy without request-time global mixing.

    Runtime contexts carry the exact admission-time object. Lower-level callers
    may instead provide a request snapshot, including an empty mapping, which is
    projected purely with deterministic defaults.
    """
    injected_export = getattr(context, "markdown_export_semantics", None)
    injected_title = getattr(context, "ocr_blockquote_title", None)
    if isinstance(injected_export, MarkdownExportSemantics):
        return MarkdownRequestPolicy(
            export=injected_export,
            ocr_blockquote_title=(
                injected_title.strip() if isinstance(injected_title, str) else _title_from_semantics(injected_export)
            ),
        )

    request = getattr(context, "request", None)
    snapshot = getattr(request, "config_snapshot", None)
    if isinstance(snapshot, Mapping):
        options = getattr(request, "options", None)
        requested_locale = options.get("locale") if isinstance(options, Mapping) else None
        export = MarkdownExportSemantics.from_config_snapshot(
            snapshot,
            requested_locale=requested_locale,
        )
        return MarkdownRequestPolicy(
            export=export,
            ocr_blockquote_title=(
                injected_title.strip() if isinstance(injected_title, str) else _title_from_semantics(export)
            ),
        )

    raise TypeError("context.request.config_snapshot must be a mapping")


# ═══════════════════════════════════════════════════════════════════════════
# Export-mode resolution
# ═══════════════════════════════════════════════════════════════════════════


def get_markdown_export_modes(
    kind: str,
    *,
    extraction_mode: str | None = None,
    ocr_placement_mode: str | None = None,
    table_merge_export_strategy: str | None = None,
    semantics: MarkdownExportSemantics,
) -> dict[str, str]:
    """Return resolved markdown export modes for a route kind.

    ``kind`` identifies the caller for API/diagnostic compatibility; it does
    not select a second configuration source. Caller overrides take precedence
    over the request's immutable, global Export-tab semantics.
    """
    _ = kind
    image_extraction_mode = str(extraction_mode or semantics.image_extraction_mode)
    resolved_ocr_placement_mode = str(ocr_placement_mode or semantics.ocr_placement_mode)
    if image_extraction_mode.strip().lower() == "base64":
        resolved_ocr_placement_mode = "main_md"

    return {
        "image_extraction_mode": image_extraction_mode,
        "ocr_placement_mode": resolved_ocr_placement_mode,
        "table_merge_export_strategy": normalize_table_merge_export_strategy(
            table_merge_export_strategy,
            default_strategy=semantics.table_merge_export_strategy,
        ),
    }


# ═══════════════════════════════════════════════════════════════════════════
# Semantic grid construction and rendering
# ═══════════════════════════════════════════════════════════════════════════


def build_table_semantic_grid(
    *,
    row_count: int,
    col_count: int,
    cell_text_by_position: Mapping[tuple[int, int], str],
    merge_regions: list[TableMergeRegion] | tuple[TableMergeRegion, ...],
) -> list[list[TableSemanticCell]]:
    """Build a 2‑D semantic cell grid from raw cell texts and merge regions.

    Each cell in the returned grid is a :class:`TableSemanticCell` that records
    its merge-region membership (if any), anchor metadata, and row/col spans.
    Non-merged cells are represented with ``is_anchor=False``, ``is_covered=False``
    and unit spans.

    F-H3b-027
    """
    # Build a lookup that maps every cell position inside a merge region
    # to its owning TableMergeRegion.
    region_lookup: dict[tuple[int, int], TableMergeRegion] = {}
    for region in merge_regions:
        for r in range(region.start_row, region.end_row + 1):
            for c in range(region.start_col, region.end_col + 1):
                region_lookup[(r, c)] = region

    grid: list[list[TableSemanticCell]] = []
    for row in range(row_count):
        row_cells: list[TableSemanticCell] = []
        for col in range(col_count):
            raw_text = str(cell_text_by_position.get((row, col), "") or "")
            region = region_lookup.get((row, col))
            if region is None:
                # Non-merged — cell stands on its own.
                row_cells.append(
                    TableSemanticCell(
                        row=row,
                        col=col,
                        raw_text=raw_text,
                        display_text=raw_text,
                        anchor_text=raw_text,
                        anchor_row=row,
                        anchor_col=col,
                        rowspan=1,
                        colspan=1,
                        is_anchor=False,
                        is_covered=False,
                    )
                )
                continue

            # This cell belongs to a merge region.
            anchor_row = region.start_row
            anchor_col = region.start_col
            anchor_text = str(cell_text_by_position.get((anchor_row, anchor_col), "") or "")
            is_anchor = row == anchor_row and col == anchor_col
            row_cells.append(
                TableSemanticCell(
                    row=row,
                    col=col,
                    raw_text=raw_text,
                    display_text=anchor_text if is_anchor else raw_text,
                    anchor_text=anchor_text,
                    anchor_row=anchor_row,
                    anchor_col=anchor_col,
                    rowspan=region.rowspan,
                    colspan=region.colspan,
                    is_anchor=is_anchor,
                    is_covered=not is_anchor,
                )
            )
        grid.append(row_cells)
    return grid


def render_table_semantic_grid(
    grid: list[list[TableSemanticCell]],
    *,
    strategy: str,
) -> list[list[str]]:
    """Render a semantic cell grid to a 2‑D string table.

    The *strategy* controls how covered cells are handled:

    ``"fill"``
        Covered cells output the **anchor text** (the value from the top-left
        cell of the merge region).
    ``"marker"``
        Covered cells output ``"<"`` (same row as anchor) or ``"^"`` (different
        row).  Non-covered cells keep their ``display_text``.
    ``"empty"``
        Covered cells output an empty string.

    Unknown strategies use ``"fill"``.

    F-H3b-029
    """
    normalized = normalize_table_merge_export_strategy(strategy)
    rendered: list[list[str]] = []
    for row_cells in grid:
        rendered_row: list[str] = []
        for cell in row_cells:
            if not cell.is_covered:
                rendered_row.append(cell.display_text)
                continue
            if normalized == "fill":
                rendered_row.append(cell.anchor_text)
            elif normalized == "marker":
                rendered_row.append(cell.marker or "")
            else:
                # "empty"
                rendered_row.append("")
        rendered.append(rendered_row)
    return rendered


# ═══════════════════════════════════════════════════════════════════════════
# Utility functions
# ═══════════════════════════════════════════════════════════════════════════


def normalize_table_merge_export_strategy(
    value: Any,
    *,
    default_strategy: str = "fill",
    log_invalid: bool = False,
    category: str | None = None,
) -> str:
    """Normalize a table-merge export strategy value.

    Returns one of ``("fill", "empty", "marker")``, falling back to
    *default_strategy* when the input is invalid.
    """
    resolved_default = str(default_strategy or "fill").strip().lower() or "fill"
    if resolved_default not in VALID_TABLE_MERGE_EXPORT_STRATEGIES:
        resolved_default = "fill"

    if value is None:
        return resolved_default

    normalized = str(value).strip().lower()
    if normalized in VALID_TABLE_MERGE_EXPORT_STRATEGIES:
        return normalized

    if log_invalid:
        suffix = f" | category={category}" if category else ""
        logger.warning(
            "Invalid merge export strategy, falling back to default: %s -> %s%s",
            value,
            resolved_default,
            suffix,
        )
    return resolved_default


def normalize_markdown_break_separator(value: Any, *, default: str) -> str:
    """Normalize configured DOCX -> Markdown separator values.

    The settings UI uses ``"ignore"`` as a sentinel.  Convert that to an
    empty string so downstream renderers skip the separator instead of
    leaking the sentinel into Markdown output.
    """
    if value is None:
        return default
    candidate = str(value).strip()
    if not candidate:
        return default
    if candidate.lower() == "ignore":
        return ""
    return candidate


def format_image_link(
    alt: str,
    target: str,
    style: str = "wiki_embed",
) -> str:
    """Format an image link in the requested *style*.

    ``wiki_embed`` → ``![[target]]``
    ``wiki_link`` → ``[[target]]``
    ``markdown_embed`` → ``![alt](target)``
    ``markdown_link`` → ``[alt](target)``
    """
    normalized = _normalize_link_style(style)
    if normalized == "wiki_embed":
        return f"![[{target}]]"
    if normalized == "wiki_link":
        return f"[[{target}]]"
    if normalized == "markdown_embed":
        return f"![{alt}]({target})"
    # markdown_link
    return f"[{alt}]({target})"


def _normalize_link_style(raw: str | None) -> str:
    """Coerce a link-style string to a known value, falling back to ``"wiki_embed"``."""
    if not raw:
        return "wiki_embed"
    candidate = str(raw).strip().lower()
    if candidate in VALID_LINK_STYLES:
        return candidate
    return "wiki_embed"


# ═══════════════════════════════════════════════════════════════════════════
# Public API
# ═══════════════════════════════════════════════════════════════════════════

__all__ = [  # noqa: RUF022
    # Dataclasses
    "LinkRuntimeConfig",
    "MarkdownExportSemantics",
    "MarkdownRequestPolicy",
    "TableMergeRegion",
    "TableSemanticCell",
    # Constants
    "VALID_LINK_STYLES",
    "VALID_TABLE_MERGE_EXPORT_STRATEGIES",
    # Semantic grid
    "build_table_semantic_grid",
    "render_table_semantic_grid",
    # Resolution
    "get_markdown_export_modes",
    "resolve_markdown_request_policy",
    # Utilities
    "format_image_link",
    "normalize_markdown_break_separator",
    "normalize_table_merge_export_strategy",
]
