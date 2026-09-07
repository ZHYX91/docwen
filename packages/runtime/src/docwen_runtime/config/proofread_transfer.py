"""Validated, side-effect-free import plans for the four proofread rule files."""

from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass

from tomlkit.items import AoT

from docwen_core.toml_tools import parse_toml_text, read_toml_text, toml_value

from .validation import ConfigSemanticError, validate_config_file

RULE_SECTIONS = {
    "proofread/pairs.toml": "items",
    "proofread/symbol_map.toml": "entries",
    "proofread/typos.toml": "entries",
    "proofread/sensitive_words.toml": "entries",
}


@dataclass(frozen=True, slots=True)
class RuleChange:
    kind: str
    key: str
    current: str
    incoming: str


@dataclass(frozen=True, slots=True)
class RuleImportPlan:
    config_name: str
    source_text: str
    changes: tuple[RuleChange, ...]
    merge_text: str | None
    replace_text: str
    merge_error: str = ""


def _comment(item: object) -> str:
    trivia = getattr(item, "trivia", None)
    return str(getattr(trivia, "comment", "") or "").lstrip("#").strip()


def _comments(item: object) -> tuple[str, ...]:
    """Collect array notes without mistaking quoted '#' characters for notes."""
    serializer = getattr(item, "as_string", None)
    source = str(serializer()) if callable(serializer) else ""
    notes = [_comment(item)]
    if isinstance(item, AoT):
        notes.extend(_comment(table) for table in item)
    index = 0
    while index < len(source):
        char = source[index]
        if char in {"'", '"'}:
            delimiter = char * 3 if source.startswith(char * 3, index) else char
            index += len(delimiter)
            while index < len(source) and not source.startswith(delimiter, index):
                index += 2 if char == '"' and source[index] == "\\" else 1
            index += len(delimiter)
        elif char == "#":
            end = source.find("\n", index)
            end = len(source) if end < 0 else end
            notes.append(source[index + 1 : end].strip())
            index = end
        else:
            index += 1
    return tuple(dict.fromkeys(part.strip() for note in notes for part in note.split("; ") if part.strip()))


def _entry_notes(table: object) -> tuple[dict[str, tuple[str, ...]], tuple[str, ...]]:
    """Associate standalone comments with the following key before merging."""
    entries: dict[str, tuple[str, ...]] = {}
    pending: list[str] = []
    for key, item in getattr(getattr(table, "value", None), "body", ()):
        if key is None:
            pending.extend(_comments(item))
        else:
            entries[str(key.key)] = tuple(dict.fromkeys((*pending, *_comments(item))))
            pending.clear()
    table_notes = tuple(
        dict.fromkeys(part.strip() for note in (_comment(table), *pending) for part in note.split("; ") if part.strip())
    )
    return entries, table_notes


def plan_rule_import(config_name: str, source_text: str, imported_text: str) -> RuleImportPlan:
    """Build both choices in memory. Saving remains the caller's transaction."""
    if config_name not in RULE_SECTIONS:
        raise ValueError("Not an editable proofread rule file")
    section = RULE_SECTIONS[config_name]
    current = validate_config_file(config_name, parse_toml_text(source_text), {})
    incoming = validate_config_file(config_name, parse_toml_text(imported_text), {})
    if section not in incoming:
        raise ConfigSemanticError(f"{config_name} must contain {section}")
    current_doc = read_toml_text(source_text)
    incoming_doc = read_toml_text(imported_text)
    replacement = deepcopy(current_doc)
    replacement[section] = deepcopy(incoming_doc[section])
    merged = deepcopy(current_doc)
    changes: list[RuleChange] = []

    if section == "items":
        old_pairs = current.get(section, [])
        new_pairs = incoming[section]
        old_by_close = {pair[1]: pair[0] for pair in old_pairs}
        new_by_close = {pair[1]: pair[0] for pair in new_pairs}
        for closing, opening in new_by_close.items():
            old = old_by_close.get(closing)
            kind = "added" if old is None else "unchanged" if old == opening else "conflict"
            changes.append(RuleChange(kind, closing, f"{old} → {closing}" if old else "", f"{opening} → {closing}"))
        for closing, opening in old_by_close.items():
            if closing not in new_by_close:
                changes.append(RuleChange("existing_only", closing, f"{opening} → {closing}", ""))
        # The validator accepts both arrays and arrays of tables. Normalize
        # only the merge container so either source can be combined safely.
        notes = "; ".join(dict.fromkeys((*_comments(current_doc.get(section)), *_comments(incoming_doc[section]))))
        merged[section] = toml_value(deepcopy(old_pairs), notes)
        old_set = {tuple(pair) for pair in old_pairs}
        for pair in new_pairs:
            if tuple(pair) not in old_set:
                merged[section].append(deepcopy(pair))
                old_set.add(tuple(pair))
    else:
        old_entries = current.get(section, {})
        new_entries = incoming[section]
        old_notes, old_table_notes = _entry_notes(current_doc.get(section))
        new_notes, new_table_notes = _entry_notes(incoming_doc[section])
        if section not in merged:
            merged[section] = {}
        if new_table_notes:
            merged[section].comment("; ".join(dict.fromkeys((*old_table_notes, *new_table_notes))))
        for key, values in new_entries.items():
            old = old_entries.get(key)
            old_item = current_doc.get(section, {}).get(key)
            new_item = incoming_doc[section][key]
            old_comment = "; ".join(old_notes.get(key, _comments(old_item)))
            new_comment = "; ".join(new_notes.get(key, _comments(new_item)))
            same = old == values and old_comment == new_comment
            kind = "added" if old is None else "unchanged" if same else "conflict"
            old_display = (repr(old) if old is not None else "") + (f"\n# {old_comment}" if old_comment else "")
            new_display = repr(values) + (f"\n# {new_comment}" if new_comment else "")
            changes.append(RuleChange(kind, key, old_display, new_display))
            if old is None:
                merged[section][key] = deepcopy(new_item)
                if new_comment:
                    merged[section][key].comment(new_comment)
            elif not same:
                comments = "; ".join(dict.fromkeys((*_comments(old_item), *new_notes.get(key, _comments(new_item)))))
                result = deepcopy(old_item)
                for index, value in enumerate(values):
                    if value not in result:
                        result.append(deepcopy(new_item[index]))
                result.comment(comments)
                merged[section][key] = result
        for key, values in old_entries.items():
            if key not in new_entries:
                changes.append(RuleChange("existing_only", key, repr(values), ""))

    replace_text = str(replacement.as_string())
    validate_config_file(config_name, parse_toml_text(replace_text), {})
    merge_text = str(merged.as_string())
    merge_error = ""
    try:
        validate_config_file(config_name, parse_toml_text(merge_text), {})
    except ConfigSemanticError as error:
        merge_text = None
        merge_error = str(error)
    return RuleImportPlan(config_name, source_text, tuple(changes), merge_text, replace_text, merge_error)
