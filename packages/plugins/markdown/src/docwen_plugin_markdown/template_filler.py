"""Template filling: inject rendered body paragraphs into a template document.

The template document contains ``{{正文}}`` or ``{{body}}`` as a body
placeholder paragraph, plus optional ``{{key}}`` placeholders for YAML
field substitution.

The renderer creates paragraphs on the document but does *not* position
them. This module relocates them to the placeholder position (one-time
batch relocation), fills YAML placeholders, and applies conditional rules.
"""

from __future__ import annotations

import contextlib
import inspect
import logging
import re
from collections.abc import Iterator, Mapping
from typing import Any

logger = logging.getLogger(__name__)


def fill_template(
    doc,
    yaml_dict: dict[str, Any],
    rendered_paragraphs: list,
    placeholder_para,
    placeholder_map: dict[str, list[Any]] | None = None,
    placeholder_rules: list[Mapping[str, Any]] | None = None,
    special_placeholder_handlers: Mapping[str, Any] | None = None,
    list_separator: str = "、",
) -> None:
    """Main entry: inject body, fill YAML placeholders, apply rules.

    Args:
        doc: python-docx ``Document``.
        yaml_dict: Extracted YAML front matter (may be empty).
        rendered_paragraphs: List of python-docx ``Paragraph`` objects
            produced by the renderer.
        placeholder_para: The body placeholder paragraph (e.g.
            containing ``{{正文}}``). ``None`` signals that body
            content should be appended at the end of the document.
        placeholder_map: Optional dict of ``{key: [paragraph, ...]}`` from
            :func:`~template_utils.scan_placeholders`. If provided,
            YAML field values are substituted into matching placeholders.
        placeholder_rules: Optional rules exposed by enabled field processor
            modules. These handle old-project conditional cleanup semantics
            such as removing table rows for empty gongwen fields.
        special_placeholder_handlers: Optional special handlers exposed by
            enabled field processor modules.
        list_separator: Exact separator used for generic YAML list values.
    """
    # ── 1. Inject rendered body paragraphs ──────────────────────────────────
    if rendered_paragraphs and placeholder_para is not None:
        inject_body_paragraphs(doc, placeholder_para, rendered_paragraphs)
    elif rendered_paragraphs:
        # The renderer already appended these elements to the main body. Move
        # the complete batch before the terminal section properties without
        # using one of those new paragraphs as a relocation anchor.
        _append_rendered_batch(doc, rendered_paragraphs)

    # ── 2. Run special placeholder handlers ─────────────────────────────────
    apply_special_placeholder_handlers(
        doc,
        yaml_dict,
        special_placeholder_handlers,
        placeholder_map=placeholder_map,
    )

    # ── 3. Fill YAML placeholders ───────────────────────────────────────────
    if placeholder_map and yaml_dict:
        fill_yaml_placeholders(
            doc,
            yaml_dict,
            placeholder_map,
            list_separator=list_separator,
        )

    # ── 4. Apply rules (conditional deletes for empty fields) ───────────────
    if placeholder_map:
        apply_placeholder_rules(doc, yaml_dict, placeholder_map, placeholder_rules=placeholder_rules)


def inject_body_paragraphs(
    doc,
    placeholder_para,
    rendered_paragraphs: list,
) -> None:
    """Batch-relocate rendered paragraphs to the placeholder position.

    This is a **one-time** operation: all *rendered_paragraphs* are
    removed from their current positions in the document body and
    inserted at the placeholder paragraph's index. The placeholder
    paragraph is then deleted.

    Args:
        doc: python-docx ``Document``.
        placeholder_para: The ``Paragraph`` containing the body
            placeholder text.
        rendered_paragraphs: List of ``Paragraph`` objects to relocate.
    """
    body = doc.element.body
    placeholder_elem = placeholder_para._element
    children = list(body)
    try:
        idx = children.index(placeholder_elem)
    except ValueError:
        logger.warning("Placeholder element not found in document body")
        return

    # Collect rendered paragraph elements, removing them from current position
    rendered_elems: list = []
    for p in rendered_paragraphs:
        e = p._element
        # Element may not be a direct child (e.g. already relocated)
        with contextlib.suppress(ValueError):
            body.remove(e)
        rendered_elems.append(e)

    # Insert at placeholder position
    for offset, elem in enumerate(rendered_elems):
        body.insert(idx + offset, elem)

    # Remove the placeholder paragraph
    body.remove(placeholder_elem)


def fill_yaml_placeholders(
    doc,
    yaml_dict: dict[str, Any],
    placeholder_map: dict[str, list[Any]],
    skip_keys: set[str] | None = None,
    list_separator: str = "、",
) -> None:
    """Replace ``{{key}}`` placeholders with corresponding YAML values.

    For each key in *placeholder_map* that exists in *yaml_dict*, every
    occurrence is replaced while preserving surrounding text and runs.

    Args:
        doc: python-docx ``Document`` (unused, for signature consistency).
        yaml_dict: YAML front matter dict.
        placeholder_map: ``{key: [paragraph, ...]}`` mapping from
            :func:`~template_utils.scan_placeholders`.
    """
    skipped = skip_keys or set()
    for key, paragraphs in placeholder_map.items():
        if key in skipped:
            continue
        value = yaml_dict.get(key)
        text = _format_placeholder_value(value, list_separator=list_separator)
        if text is None:
            continue

        for para in paragraphs:
            _replace_placeholder_runs(para, key, text)


def _replace_placeholder_runs(para, key: str, replacement: str) -> None:
    """Replace all occurrences without flattening runs or hyperlinks."""
    pattern = re.compile(r"\{\{\s*" + re.escape(key) + r"\s*\}\}")
    characters: list[str] = []
    locations: list[tuple[Any, int] | None] = []
    for element in para._element.iter():
        if not isinstance(element.tag, str):
            continue
        tag_name = _local_name(element.tag)
        if tag_name == "t":
            for offset, character in enumerate(element.text or ""):
                characters.append(character)
                locations.append((element, offset))
        elif tag_name in {"br", "cr", "tab"}:
            # A placeholder must not be stitched across a visible structural
            # boundary such as a tab or line break.
            characters.append("\0")
            locations.append(None)

    matches = list(pattern.finditer("".join(characters)))

    for match in reversed(matches):
        start_location = locations[match.start()]
        end_location = locations[match.end() - 1]
        if start_location is None or end_location is None:
            continue
        start_element, start_offset = start_location
        end_element, end_offset = end_location
        end_offset += 1
        if start_element is end_element:
            current_text = start_element.text or ""
            _set_word_text(start_element, current_text[:start_offset] + replacement + current_text[end_offset:])
            continue
        start_text = start_element.text or ""
        end_text = end_element.text or ""
        _set_word_text(start_element, start_text[:start_offset] + replacement)
        seen_elements: set[Any] = {start_element, end_element}
        for location in locations[match.start() + 1 : match.end() - 1]:
            if location is None:
                continue
            element, _offset = location
            if element not in seen_elements:
                _set_word_text(element, "")
                seen_elements.add(element)
        _set_word_text(end_element, end_text[end_offset:])


def _set_word_text(element, text: str) -> None:
    element.text = text
    xml_space = "{http://www.w3.org/XML/1998/namespace}space"
    if text[:1].isspace() or text[-1:].isspace():
        element.set(xml_space, "preserve")
    else:
        element.attrib.pop(xml_space, None)


def apply_special_placeholder_handlers(
    doc,
    yaml_dict: dict[str, Any],
    special_placeholder_handlers: Mapping[str, Any] | None = None,
    *,
    placeholder_map: Mapping[str, list[Any]] | None = None,
) -> set[str]:
    """Run special handlers only for placeholders from the template snapshot."""
    handled: set[str] = set()
    for key, handler in (special_placeholder_handlers or {}).items():
        if not callable(handler):
            continue
        paragraphs = None if placeholder_map is None else list(placeholder_map.get(key, ()))
        if placeholder_map is not None and not paragraphs:
            continue
        if paragraphs is not None and _accepts_placeholder_paragraphs(handler):
            result = handler(doc, yaml_dict, placeholder_paragraphs=paragraphs)
        else:
            result = handler(doc, yaml_dict)
        if result:
            handled.add(key)
    return handled


def _accepts_placeholder_paragraphs(handler: Any) -> bool:
    try:
        parameters = inspect.signature(handler).parameters.values()
    except (TypeError, ValueError):
        return False
    return any(
        parameter.kind is inspect.Parameter.VAR_KEYWORD or parameter.name == "placeholder_paragraphs"
        for parameter in parameters
    )


def apply_placeholder_rules(
    doc,
    yaml_dict: dict[str, Any],
    placeholder_map: dict[str, list[Any]],
    placeholder_rules: list[Mapping[str, Any]] | None = None,
) -> None:
    """Apply configured empty-field cleanup and replace remaining empties.

    Current rules:
    - Enabled field processor modules may expose old-project cleanup rules,
      such as removing a row that contains a complete group of empty gongwen
      placeholders.
    - Missing or empty fields that do not trigger a configured container rule
      are replaced with an empty string. Their surrounding paragraph, row, or
      cell is retained.

    Args:
        doc: python-docx ``Document``.
        yaml_dict: YAML front matter dict.
        placeholder_map: ``{key: [paragraph, ...]}`` mapping.
        placeholder_rules: Optional field-processor cleanup rules.
    """
    rule_sets = list(placeholder_rules or [])
    operations = (
        ("delete_table_if_empty", _remove_tables_for_fields),
        ("delete_row_if_empty", _remove_table_rows_for_fields),
        ("delete_paragraph_if_empty", _remove_body_paragraphs_for_fields),
        ("delete_cell_if_empty", _clear_table_cells_for_fields),
    )
    # Apply one operation class across every processor before advancing to the
    # next class. This preserves table/row priority even when multiple field
    # processor modules contribute rules in different mapping orders.
    for rule_name, operation in operations:
        for rules in rule_sets:
            for field_group in _iter_field_groups(rules.get(rule_name)):
                if _all_fields_empty(yaml_dict, field_group):
                    operation(doc, placeholder_map, field_group)

    for key, paragraphs in placeholder_map.items():
        value = yaml_dict.get(key)
        if not _is_empty_placeholder_value(value):
            continue
        # A configured rule may already have removed the owning container. Any
        # remaining occurrence follows the historical default: replace just
        # the placeholder and preserve surrounding content and formatting.
        for para in paragraphs:
            _replace_placeholder_runs(para, key, "")


def _append_rendered_batch(doc, paragraphs: list) -> None:
    """Keep one rendered batch ordered immediately before terminal ``sectPr``."""

    body = doc.element.body
    elements: list = []
    for paragraph in paragraphs:
        element = paragraph._element
        with contextlib.suppress(ValueError):
            body.remove(element)
        elements.append(element)
    terminal_section = body.sectPr
    insertion_index = body.index(terminal_section) if terminal_section is not None else len(body)
    for offset, element in enumerate(elements):
        body.insert(insertion_index + offset, element)


def _format_placeholder_value(value: Any, *, list_separator: str = "、") -> str | None:
    if _is_empty_placeholder_value(value):
        return None
    if isinstance(value, (list, tuple)):
        items = list(_iter_placeholder_list_items(value))
        return list_separator.join(items)
    return str(value)


def _iter_placeholder_list_items(value: list[Any] | tuple[Any, ...]) -> Iterator[str]:
    for item in value:
        if isinstance(item, (list, tuple)):
            yield from _iter_placeholder_list_items(item)
            continue
        if _is_empty_placeholder_value(item):
            continue
        if isinstance(item, str) and item in {"null", "None"}:
            continue
        yield str(item)


def _is_empty_placeholder_value(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == ""
    if isinstance(value, Mapping):
        return all(_is_empty_placeholder_value(item) for item in value.values())
    if isinstance(value, (list, tuple, set)):
        return all(_is_empty_placeholder_value(item) for item in value)
    return False


def _iter_field_groups(value: Any) -> list[list[str]]:
    if not isinstance(value, list):
        return []
    groups: list[list[str]] = []
    for item in value:
        if isinstance(item, list):
            fields = [field for field in item if isinstance(field, str) and field]
        elif isinstance(item, str) and item:
            fields = [item]
        else:
            fields = []
        if fields:
            groups.append(fields)
    return groups


def _all_fields_empty(yaml_dict: dict[str, Any], field_names: list[str]) -> bool:
    return all(_is_empty_placeholder_value(yaml_dict.get(field_name)) for field_name in field_names)


def _remove_body_paragraphs_for_fields(
    doc,
    placeholder_map: Mapping[str, list[Any]],
    field_names: list[str],
) -> None:
    body = doc.element.body
    for paragraph in _common_placeholder_containers(placeholder_map, field_names, "p"):
        if paragraph.getparent() is body:
            body.remove(paragraph)


def _remove_table_rows_for_fields(
    _doc,
    placeholder_map: Mapping[str, list[Any]],
    field_names: list[str],
) -> None:
    affected_tables: set[Any] = set()
    for row in _common_placeholder_containers(placeholder_map, field_names, "tr"):
        table = row.getparent()
        if table is None or _local_name(table.tag) != "tbl":
            continue
        table.remove(row)
        affected_tables.add(table)
    for table in affected_tables:
        if not any(isinstance(child.tag, str) and _local_name(child.tag) == "tr" for child in table):
            parent = table.getparent()
            if parent is not None:
                parent.remove(table)


def _clear_table_cells_for_fields(
    _doc,
    placeholder_map: Mapping[str, list[Any]],
    field_names: list[str],
) -> None:
    for cell in _common_placeholder_containers(placeholder_map, field_names, "tc"):
        if _nearest_ancestor_with_tag(cell, "tbl") is None:
            continue
        for child in cell.iter():
            if (
                isinstance(child.tag, str)
                and _local_name(child.tag) == "t"
                and _nearest_ancestor_with_tag(child, "tc") is cell
            ):
                _set_word_text(child, "")


def _remove_tables_for_fields(
    _doc,
    placeholder_map: Mapping[str, list[Any]],
    field_names: list[str],
) -> None:
    for table in _common_placeholder_containers(placeholder_map, field_names, "tbl"):
        parent = table.getparent()
        if parent is not None:
            parent.remove(table)


def _common_placeholder_containers(
    placeholder_map: Mapping[str, list[Any]],
    field_names: list[str],
    container_tag: str,
) -> set[Any]:
    common: set[Any] | None = None
    for field_name in field_names:
        containers: set[Any] = set()
        for paragraph in placeholder_map.get(field_name, ()):
            paragraph_element = paragraph._element
            container = (
                paragraph_element
                if container_tag == "p"
                else _nearest_ancestor_with_tag(paragraph_element, container_tag)
            )
            if container is not None:
                containers.add(container)
        common = containers if common is None else common & containers
        if not common:
            return set()
    return common or set()


def _nearest_ancestor_with_tag(element, tag_name: str):
    current = element.getparent()
    while current is not None:
        if isinstance(current.tag, str) and _local_name(current.tag) == tag_name:
            return current
        current = current.getparent()
    return None


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]
