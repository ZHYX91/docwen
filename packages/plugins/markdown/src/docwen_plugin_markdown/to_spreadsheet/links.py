"""Request-owned link drafts; only an entire single-link cell is clickable."""

from __future__ import annotations

import re
import secrets
from dataclasses import replace
from pathlib import Path
from typing import Any

from docwen_core.export_semantics import LinkRuntimeConfig
from docwen_core.links import process_markdown_links, project_field_links
from docwen_plugin_markdown.mistune_extensions import parse_markdown_text
from docwen_plugin_markdown.renderer_inlines import _resolve_hyperlink_target, extract_text_content, plain_link_label
from docwen_plugin_markdown.xlsx_text import assign_bounded_xlsx_value, existing_xlsx_cells


class SpreadsheetLinks:
    def __init__(self, source: str, config: LinkRuntimeConfig, *, native: bool) -> None:
        self.source = source
        self.config = config
        self.native = native
        self.prefix = "DOCWENSPREADSHEETLINK" + secrets.token_hex(24)
        self.links: dict[str, tuple[str, str, str]] = {}

    def issue(self, label: str, target: str, original: str, syntax: str) -> str:
        nodes = parse_markdown_text(label, consume_initial_bom=False)
        visible = "".join(extract_text_content(node.get("children", [])) for node in nodes)
        mode = self.config.wiki_hyperlink_fallback if syntax == "wiki" else self.config.markdown_hyperlink_fallback
        # Run extraction on the original source, never the resolved absolute URI.
        fallback_config = replace(
            self.config,
            non_embed_wiki_mode=mode if syntax == "wiki" else "keep",
            non_embed_markdown_mode=mode if syntax == "markdown" else "keep",
        )
        fallback = process_markdown_links(original, self.source, link_config=fallback_config, target_format="csv")
        if not visible and not label:
            address_config = replace(
                fallback_config, non_embed_wiki_mode="extract_address", non_embed_markdown_mode="extract_address"
            )
            visible = process_markdown_links(original, self.source, link_config=address_config, target_format="csv")
        if mode == "extract_text":
            fallback = visible or label
        token = f"{self.prefix}N{len(self.links)}END"
        resolved = _resolve_hyperlink_target(target, Path(self.source).parent)
        if any(ord(char) < 32 for char in resolved) or not resolved.lower().startswith(
            ("http://", "https://", "mailto:", "ftp://", "file://", "obsidian://")
        ):
            resolved = ""
        self.links[token] = (visible or label, resolved, fallback)
        return token

    def issue_literal(self, value: str) -> str:
        token = f"{self.prefix}N{len(self.links)}END"
        self.links[token] = (value, "", value)
        return token

    def restore(self, value: Any) -> Any:
        if not isinstance(value, str):
            return value
        pattern = re.compile(re.escape(self.prefix) + r"N\d+END")
        return pattern.sub(
            lambda match: self.links[match.group()][2] if match.group() in self.links else match.group(), value
        )

    def project(self, value: Any) -> Any:
        if isinstance(value, dict):
            return {key: self.project(item) for key, item in value.items()}
        if isinstance(value, (list, tuple)):
            return type(value)(self.project(item) for item in value)
        if not isinstance(value, str):
            return value
        # Field processing must not resolve or expand any embedded resource.
        return project_field_links(
            value,
            source_file_path=self.source,
            wiki_mode=self.config.non_embed_wiki_mode,
            markdown_mode=self.config.non_embed_markdown_mode,
            search_dirs=self.config.search_dirs,
            on_not_found=self.config.file_not_found_mode,
            wiki_hyperlink_fallback=self.config.wiki_hyperlink_fallback,
            markdown_hyperlink_fallback=self.config.markdown_hyperlink_fallback,
            target_format="xlsx" if self.native else "csv",
            hyperlink_renderer=lambda label, target: label,
            hyperlink_projection=self.issue if self.native else None,
            literal_projection=self.issue_literal,
            markdown_text_renderer=plain_link_label,
        )

    def write_cell(self, cell: Any, value: Any) -> None:
        if not isinstance(value, str) or self.prefix not in value:
            assign_bounded_xlsx_value(cell, value)
            return
        link = self.links.get(value)
        if self.native and link is not None and link[1]:
            label, target, _ = link
            assign_bounded_xlsx_value(cell, label)
            cell.data_type = "s"
            cell.hyperlink = target
        else:
            assign_bounded_xlsx_value(cell, self.restore(value))
            cell.data_type = "s"
            cell.hyperlink = None

    def materialize(self, workbook: Any) -> None:
        for sheet in workbook.worksheets:
            for cell in existing_xlsx_cells(sheet):
                if isinstance(cell.value, str) and self.prefix in cell.value:
                    self.write_cell(cell, cell.value)
