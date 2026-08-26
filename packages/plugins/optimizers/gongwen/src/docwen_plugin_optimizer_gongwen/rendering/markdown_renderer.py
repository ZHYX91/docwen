"""Generate Markdown output with YAML frontmatter for gongwen documents."""

from __future__ import annotations

from collections.abc import Mapping
from pathlib import Path, PureWindowsPath
from typing import TYPE_CHECKING, Any

import yaml

from docwen_core.text.image_markdown import generate_image_markdown
from docwen_core.text.splitting import split_once

if TYPE_CHECKING:
    from docwen_core.export_semantics import MarkdownExportSemantics
    from docwen_plugin_optimizer_gongwen.models import GongwenMetadata, ParagraphFeature
    from docwen_plugin_optimizer_gongwen.runtime_config import GongwenContentRuntimeConfig


def render(
    metadata: GongwenMetadata,
    body_lines: list[str],
    skip_indices: list[int] | None = None,
    *,
    feature_map: Mapping[int, ParagraphFeature],
    config: GongwenContentRuntimeConfig | None = None,
    remove_numbering: bool = True,
    heading_formatter: Any | None = None,
    image_mode: str = "file",
    image_link_style: str = "markdown_embed",
    export_semantics: MarkdownExportSemantics | None = None,
) -> str:
    """Render gongwen metadata + body paragraphs as Markdown.

    Args:
        metadata: The 18-field gongwen metadata.
        body_lines: Text content of all paragraphs.
        skip_indices: Indices of paragraphs to skip (structural elements).
        feature_map: Explicit mapping from paragraph index to ParagraphFeature
            for rich content rendering (formulas, images, headings, breaks).
        config: Optional runtime configuration.
        remove_numbering: Whether to strip the detected heading numbering
            prefix from heading output (default True).
        heading_formatter: Optional ``HeadingFormatter`` instance to apply
            scheme-based numbering. When provided, the cleaned heading text
            is passed through the formatter.

    Returns:
        Markdown string with optional YAML frontmatter.
    """
    output_parts: list[str] = []

    # Build YAML frontmatter if metadata has any non-empty fields
    yaml_data = _build_yaml_data(metadata)
    if yaml_data:
        output_parts.append("---")
        output_parts.append(
            yaml.dump(
                yaml_data,
                allow_unicode=True,
                default_flow_style=False,
                sort_keys=False,
            ).rstrip()
        )
        output_parts.append("---")
        output_parts.append("")

    # Add body text with rich content rendering
    skip = set(skip_indices or [])
    fm = feature_map

    def append_images(pf: ParagraphFeature) -> None:
        for img_path in pf.extracted_images:
            if image_mode == "base64":
                markdown_path = img_path
            elif "\\" in img_path:
                markdown_path = PureWindowsPath(img_path).name
            else:
                markdown_path = Path(img_path).name
            img_md = generate_image_markdown(
                image_path=markdown_path,
                image_mode=image_mode,
                image_link_style=image_link_style,
                export_semantics=export_semantics,
            )
            ocr_text = pf.image_ocr_texts.get(img_path, "")
            if ocr_text:
                img_md += f"\n> {ocr_text}"
            output_parts.append(img_md)
            output_parts.append("")

    for i, line in enumerate(body_lines):
        pf = fm.get(i)
        if i in skip:
            # Structural text may move into YAML, but a seal/signature image in
            # the same paragraph must remain reachable in the Markdown output.
            if pf and pf.extracted_images:
                append_images(pf)
            continue

        # Allow empty-text paragraphs that have rich content (formulas, images, breaks)
        has_rich = pf and (pf.has_formula or pf.has_page_break or pf.has_section_break or pf.extracted_images)
        if not line.strip() and not has_rich:
            continue

        # Page/section break before paragraph (respect runtime config)
        if pf and pf.has_page_break:
            marker = config.page_break_marker if config and config.horizontal_rule_enabled else "---"
            output_parts.append(marker)
            output_parts.append("")
        if pf and pf.has_section_break:
            marker = config.section_break_marker if config and config.horizontal_rule_enabled else "---"
            output_parts.append(marker)
            output_parts.append("")

        if pf and pf.is_table_anchor and pf.table_markdown:
            output_parts.extend(pf.table_markdown.splitlines())
            output_parts.append("")
            continue

        # Formula rendering
        if pf and pf.has_formula and pf.formula_latex:
            if pf.formula_type == "block":
                output_parts.append(f"$$\n{pf.formula_latex}\n$$")
            else:
                # Inline formula: include text and LaTeX
                output_parts.append(f"{line} ${pf.formula_latex}$")
            output_parts.append("")
            # Images after formula paragraph
            if pf.extracted_images:
                append_images(pf)
            continue

        # Heading rendering
        if pf and pf.heading_level > 0:
            # Determine original numbering prefix
            numbering = "" if remove_numbering else pf.heading_numbering_text or ""

            # Extraction owns the Gongwen-specific boundary decision. Rendering
            # only applies the explicit character boundary to this same source
            # feature so later images/formulas/skip flags cannot shift indices.
            if pf.heading_body_boundary is not None:
                heading_content, mixed_body_content = split_once(line, pf.heading_body_boundary)
                heading_content = heading_content.strip()
                mixed_body_content = mixed_body_content.strip()
            else:
                heading_content, mixed_body_content = line, ""

            # Add new scheme-based numbering if formatter is provided
            if heading_formatter is not None:
                heading_content = heading_formatter.format_heading(heading_content, pf.heading_level)

            heading_text = f"{'#' * pf.heading_level} {numbering}{heading_content}".rstrip()
            output_parts.append(heading_text)
            if mixed_body_content:
                output_parts.append(mixed_body_content)
            output_parts.append("")
            # Images after heading paragraph
            if pf.extracted_images:
                append_images(pf)
            continue

        # Plain text
        output_parts.append(line)

        # Images after paragraph
        if pf and pf.extracted_images:
            append_images(pf)

        output_parts.append("")

    return "\n".join(output_parts)


def _build_yaml_data(metadata: GongwenMetadata) -> dict:
    """Build the fixed 18-field YAML schema when any metadata is present."""
    d = metadata.to_dict()
    has_metadata = any(value and not (isinstance(value, list) and len(value) == 0) for value in d.values())
    return d if has_metadata else {}
