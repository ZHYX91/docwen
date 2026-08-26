"""YAML frontmatter builder and Markdown table renderer for invoice output."""

from __future__ import annotations


def _format_yaml_value(value: str) -> str:
    """Format a string value for safe YAML inclusion.

    Returns the value quoted if it contains characters that need quoting,
    otherwise returns it bare.
    """
    v = str(value).strip()
    if not v:
        return "''"
    # Characters that require quoting in YAML
    if any(
        ch in v
        for ch in (
            ":",
            "#",
            "{",
            "}",
            "[",
            "]",
            ",",
            "&",
            "*",
            "?",
            "|",
            "-",
            "<",
            ">",
            "=",
            "!",
            "%",
            "@",
            "`",
            "'",
            '"',
        )
    ):
        # Use single-quoted style, escaping internal single quotes
        escaped = v.replace("'", "''")
        return f"'{escaped}'"
    if v.lower() in ("true", "false", "yes", "no", "null", "~"):
        return f"'{v}'"
    if v[0] in "0123456789" and any(ch in v for ch in ".:"):
        return f"'{v}'"
    return v


def _build_yaml_frontmatter(
    *,
    file_stem: str,
    metadata: dict[str, str | None],
    include_empty: bool = False,
    yaml_key_labels: dict[str, str] | None = None,
) -> str:
    """Build a YAML frontmatter block from invoice metadata.

    Args:
        file_stem: The file stem to use for the ``aliases`` and ``title`` fields.
        metadata: Invoice metadata dict (field name → value).
        include_empty: If True, include fields with empty values as ``key: ''``.
        yaml_key_labels: Locale-resolved labels supplied by the app boundary.

    Returns:
        A string containing the YAML frontmatter block (with ``---`` delimiters).
    """
    lines: list[str] = ["---"]

    safe_stem = _format_yaml_value(file_stem)
    lines.append("aliases:")
    lines.append(f"  - {safe_stem}")

    title_key = "标题"
    if isinstance(yaml_key_labels, dict):
        title_label = yaml_key_labels.get("title")
        if isinstance(title_label, str) and title_label.strip():
            title_key = title_label.strip()
    lines.append(f"{title_key}: {safe_stem}")

    for key, value in metadata.items():
        if value is None:
            if include_empty:
                lines.append(f"{key}: ''")
            continue
        value = str(value).strip()
        if not value:
            if include_empty:
                lines.append(f"{key}: ''")
            continue
        lines.append(f"{key}: {_format_yaml_value(value)}")

    lines.append("---")
    lines.append("")
    return "\n".join(lines)


def _escape_table_cell(value: str) -> str:
    """Escape pipe and newline characters in a Markdown table cell."""
    return str(value).replace("|", "\\|").replace("\n", " ").strip()


def _render_markdown_table(*, headers: list[str], rows: list[dict[str, str]]) -> str:
    """Render a Markdown table from headers and row dicts.

    Args:
        headers: Column header names.
        rows: List of row dicts (column name → value).

    Returns:
        A string containing the Markdown table.
    """
    normalized_rows: list[list[str]] = []
    for row in rows:
        normalized_rows.append([_escape_table_cell(row.get(h, "")) for h in headers])

    if not normalized_rows:
        normalized_rows = [[_escape_table_cell("（未识别）")] + [""] * (len(headers) - 1)]

    out: list[str] = []
    out.append("| " + " | ".join(headers) + " |")
    out.append("| " + " | ".join(["---"] * len(headers)) + " |")
    for r in normalized_rows:
        out.append("| " + " | ".join(r) + " |")
    out.append("")
    return "\n".join(out)
