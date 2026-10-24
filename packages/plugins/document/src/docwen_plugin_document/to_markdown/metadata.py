"""Word title/subtitle semantics and localized Markdown front matter.

These pure helpers have no conversion state, I/O or runtime-resource access.
The converter owns policy selection and output publication.
"""

from __future__ import annotations

from typing import Any


def format_yaml_value(value: Any) -> str:
    """Format a value as a YAML-safe string.

    Wraps values in single quotes when they contain characters that
    YAML parsers would misinterpret (leading ``[``, ``{``, ``#``,
    ``&``, ``*``, ``!``, ``|``, ``>``, ``%``, ``@``, ````` ``,
    ``'``, ``"``, ``:``, or trailing ``#`` with preceding space).
    Handles None/empty gracefully.

    Used by ``build_yaml_header`` for safe YAML front matter.
    """
    if value is None:
        return ""
    s = str(value).replace("\n", " ").replace("\r", "")
    s = s.strip()
    if not s:
        return ""

    special_lead = {"[", "{", "#", "&", "*", "!", "|", ">", "%", "@", "`"}
    needs_quote = (
        s[0] in special_lead
        or (len(s) > 1 and s[0] == "-" and (s[1].isspace() or s[1] == " "))
        or "'" in s
        or '"' in s
        or ": " in s
        or s.endswith(":")
        or " #" in s
        or s.lower() in ("true", "false", "null", "yes", "no", "on", "off", "~")
        or s.isdigit()
        or (s.startswith("0") and len(s) > 1 and s[1].isdigit())
    )

    if needs_quote:
        if '"' not in s:
            return f'"{s}"'
        escaped = s.replace("'", "''")
        return f"'{escaped}'"
    return s


def build_yaml_header(
    metadata: dict,
    locale: str = "en",
    yaml_key_labels: object | None = None,
) -> str:
    """Build a YAML front matter string from metadata dict.

    Produces a complete YAML front matter block (``---`` delimiters,
    ``aliases:`` list, ``title:`` / ``subtitle:`` fields, closing
    ``---``) suitable for prepending to Markdown output.

    Args:
        metadata: Dict with keys ``aliases`` (list), ``title`` (str),
            ``subtitle`` (str).
        locale: Locale string for localized YAML key names.
        yaml_key_labels: Optional pre-resolved labels from the application
            edge, e.g. ``{"title": "Titel", "subtitle": "Untertitel"}``.

    Returns:
        A YAML front matter string ending with ``\\n``, or an empty
        string if the metadata is empty/trivial.
    """
    aliases = metadata.get("aliases") or []
    title = metadata.get("title") or ""
    subtitle = metadata.get("subtitle") or ""

    parts: list[str] = ["---"]

    # L3: Locale-aware YAML key names.  Plugins consume resolved labels
    # instead of loading runtime resources themselves.
    if locale in ("zh_CN", "zh-CN", "zh"):
        key_aliases = "别名"
        key_title = "标题"
        key_subtitle = "副标题"
    else:
        key_aliases = "aliases"
        key_title = "title"
        key_subtitle = "subtitle"
    if isinstance(yaml_key_labels, dict):
        label_title = yaml_key_labels.get("title")
        label_subtitle = yaml_key_labels.get("subtitle")
        label_aliases = yaml_key_labels.get("aliases")
        if isinstance(label_title, str) and label_title.strip():
            key_title = label_title.strip()
        if isinstance(label_subtitle, str) and label_subtitle.strip():
            key_subtitle = label_subtitle.strip()
        if isinstance(label_aliases, str) and label_aliases.strip():
            key_aliases = label_aliases.strip()

    if aliases:
        parts.append(f"{key_aliases}:")
        for alias in aliases:
            safe = format_yaml_value(alias)
            parts.append(f"  - {safe}")
    else:
        parts.append(f"{key_aliases}: []")

    safe_title = format_yaml_value(title)
    parts.append(f"{key_title}: {safe_title}" if safe_title else f"{key_title}: ")

    safe_subtitle = format_yaml_value(subtitle)
    parts.append(f"{key_subtitle}: {safe_subtitle}" if safe_subtitle else f"{key_subtitle}: ")

    parts.append("---")
    parts.append("")
    return "\n".join(parts) + "\n"


def extract_title_metadata(
    doc: Any,
    docx_path: str,
) -> tuple[dict, set]:
    """Extract title/subtitle metadata from a python-docx Document.

    Scans paragraphs for Word "Title" and "Subtitle" styles at the
    start of the document.  Collects consecutive matches, merges
    multi-paragraph titles, and falls back to the filename stem
    when no Title style is found.

    Args:
        doc: A python-docx ``Document`` object.
        docx_path: Absolute path to the DOCX file (for filename
            fallback).

    Returns:
        A tuple of ``(metadata, skip_indices)`` where ``metadata``
        is a dict with keys ``aliases``, ``title``, ``subtitle``
        and ``skip_indices`` is a set of paragraph indices to skip
        in the body rendering.
    """
    import os as _os

    metadata: dict = {"aliases": [], "title": "", "subtitle": ""}
    title_indices: list[int] = []
    subtitle_indices: list[int] = []
    started_body = False  # True once we have seen a non-title/subtitle paragraph

    for i, para in enumerate(doc.paragraphs):
        style_name = para.style.name if para.style else ""

        if started_body:
            break  # Title/Subtitle only meaningful at document start

        if style_name == "Title":
            title_indices.append(i)
        elif style_name == "Subtitle":
            subtitle_indices.append(i)
        else:
            # First non-title, non-subtitle paragraph starts body content
            started_body = True
            # Do NOT break yet — we still need to process the accumulated
            # title/subtitle indices, which is done below the loop.

    # Merge title paragraphs
    if title_indices:
        title = "".join(doc.paragraphs[i].text.replace("\n", "").replace("\r", "").strip() for i in title_indices)
        metadata["title"] = title
        metadata["aliases"].append(title)
    else:
        filename = _os.path.basename(docx_path)
        stem = filename.rsplit(".", 1)[0]
        metadata["title"] = stem
        metadata["aliases"].append(stem)

    # Merge subtitle paragraphs
    if subtitle_indices:
        subtitle = "".join(doc.paragraphs[i].text.replace("\n", "").replace("\r", "").strip() for i in subtitle_indices)
        metadata["subtitle"] = subtitle

    skip_indices: set = set(title_indices) | set(subtitle_indices)
    return metadata, skip_indices
