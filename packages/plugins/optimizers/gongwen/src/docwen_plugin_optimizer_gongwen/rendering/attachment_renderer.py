"""Generate Markdown for attachment sections."""

from __future__ import annotations

from datetime import UTC, datetime
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from docwen_plugin_optimizer_gongwen.models import GongwenMetadata


def render_attachment(
    metadata: GongwenMetadata,
    attachment_lines: list[str],
    input_path: str = "",
) -> str:
    """Render attachment content as Markdown with YAML frontmatter header.

    Args:
        metadata: Gongwen metadata with attachment field.
        attachment_lines: Text lines of attachment paragraphs.
        input_path: Path of the source DOCX file (for YAML header).

    Returns:
        Markdown string for the attachment section, or empty string if no content.
    """
    if not attachment_lines:
        return ""

    parts: list[str] = []

    # ── YAML frontmatter header ──
    source_name = Path(input_path).name if input_path else "unknown.docx"
    timestamp = datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%S")
    parts.append("---")
    parts.append(f"来源文件: {source_name}")
    parts.append(f"提取时间: {timestamp}")
    parts.append("类型: 附件内容")
    parts.append("---")
    parts.append("")

    # Add attachment header if present in metadata
    if metadata.attachment:
        parts.append("## 附件")
        parts.append("")
        for i, att in enumerate(metadata.attachment, 1):
            parts.append(f"{i}. {att}")
        parts.append("")

    # Add attachment body
    for line in attachment_lines:
        if line.strip():
            parts.append(line)
            parts.append("")

    return "\n".join(parts)
