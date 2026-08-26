"""Closed wire contract for already-presented bibliography resources."""

from __future__ import annotations

import json
from collections.abc import Mapping
from dataclasses import dataclass
from typing import Any, NoReturn

from docwen_core.models.semantic_document import (
    SemanticBibliographyEntry,
    SemanticBibliographyFragment,
    SemanticBibliographyRun,
    SemanticDocument,
    validate_semantic_document,
)

SEMANTIC_BIBLIOGRAPHY_SCHEMA = "docwen.semantic_bibliography.v1"
SEMANTIC_BIBLIOGRAPHY_MEDIA_TYPE = "application/vnd.docwen.semantic-bibliography+json"
SEMANTIC_BIBLIOGRAPHY_MAX_BYTES = 8 * 1024 * 1024


@dataclass(frozen=True, slots=True)
class SemanticBibliographyResourceError(ValueError):
    """Stable fail-closed resource error."""

    code: str
    message: str

    def __str__(self) -> str:
        return self.message


def parse_semantic_bibliography(data: bytes) -> SemanticBibliographyFragment:
    """Parse one closed v1 resource without accepting JSON extensions."""

    if len(data) > SEMANTIC_BIBLIOGRAPHY_MAX_BYTES:
        _fail("semantic_bibliography.resource_too_large", "Bibliography resource exceeds the 8 MiB limit.")
    try:
        text = data.decode("utf-8", errors="strict")
    except UnicodeDecodeError:
        _fail("semantic_bibliography.invalid_utf8", "Bibliography resource must be valid UTF-8.")
    try:
        payload = json.loads(
            text,
            object_pairs_hook=_closed_object,
            parse_constant=_reject_nonfinite,
        )
    except SemanticBibliographyResourceError:
        raise
    except (json.JSONDecodeError, RecursionError):
        _fail("semantic_bibliography.invalid_json", "Bibliography resource must be one valid JSON value.")

    root = _require_object(payload, "root")
    _require_keys(root, required={"schema", "entries"}, allowed={"schema", "entries"}, location="root")
    if root["schema"] != SEMANTIC_BIBLIOGRAPHY_SCHEMA:
        _fail(
            "semantic_bibliography.schema_unsupported",
            f"Bibliography schema must be {SEMANTIC_BIBLIOGRAPHY_SCHEMA!r}.",
        )
    entries_value = root["entries"]
    if not isinstance(entries_value, list):
        _fail("semantic_bibliography.shape_invalid", "Bibliography entries must be an array.")

    entries: list[SemanticBibliographyEntry] = []
    for entry_index, raw_entry in enumerate(entries_value):
        location = f"entries[{entry_index}]"
        entry = _require_object(raw_entry, location)
        _require_keys(entry, required={"item_id", "runs"}, allowed={"item_id", "runs"}, location=location)
        item_id = entry["item_id"]
        if not isinstance(item_id, str):
            _fail("semantic_bibliography.shape_invalid", f"{location}.item_id must be a string.")
        runs_value = entry["runs"]
        if not isinstance(runs_value, list):
            _fail("semantic_bibliography.shape_invalid", f"{location}.runs must be an array.")
        runs: list[SemanticBibliographyRun] = []
        for run_index, raw_run in enumerate(runs_value):
            run_location = f"{location}.runs[{run_index}]"
            run = _require_object(raw_run, run_location)
            _require_keys(
                run,
                required={"text"},
                allowed={"text", "bold", "italic", "href"},
                location=run_location,
            )
            text_value = run["text"]
            bold_value = run.get("bold", False)
            italic_value = run.get("italic", False)
            href_value = run.get("href")
            if not isinstance(text_value, str):
                _fail("semantic_bibliography.shape_invalid", f"{run_location}.text must be a string.")
            if type(bold_value) is not bool:
                _fail("semantic_bibliography.shape_invalid", f"{run_location}.bold must be a boolean.")
            if type(italic_value) is not bool:
                _fail("semantic_bibliography.shape_invalid", f"{run_location}.italic must be a boolean.")
            if "href" in run and not isinstance(href_value, str):
                _fail("semantic_bibliography.shape_invalid", f"{run_location}.href must be a string.")
            runs.append(
                SemanticBibliographyRun(
                    text=text_value,
                    bold=bold_value,
                    italic=italic_value,
                    href=href_value,
                )
            )
        entries.append(SemanticBibliographyEntry(item_id=item_id, runs=tuple(runs)))

    fragment = SemanticBibliographyFragment(entries=tuple(entries))
    diagnostics = validate_semantic_document(SemanticDocument(blocks=(), bibliography=fragment))
    if diagnostics:
        _fail("semantic_bibliography.content_invalid", diagnostics[0].message)
    return fragment


def _closed_object(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    value: dict[str, Any] = {}
    for key, item in pairs:
        if key in value:
            _fail("semantic_bibliography.duplicate_key", f"Duplicate JSON key {key!r} is not allowed.")
        value[key] = item
    return value


def _reject_nonfinite(value: str) -> NoReturn:
    _fail("semantic_bibliography.nonfinite_number", f"Non-finite JSON number {value!r} is not allowed.")


def _require_object(value: object, location: str) -> Mapping[str, Any]:
    if not isinstance(value, dict):
        _fail("semantic_bibliography.shape_invalid", f"Bibliography {location} must be an object.")
    return value


def _require_keys(
    value: Mapping[str, Any],
    *,
    required: set[str],
    allowed: set[str],
    location: str,
) -> None:
    missing = required - set(value)
    unknown = set(value) - allowed
    if missing:
        _fail("semantic_bibliography.shape_invalid", f"Bibliography {location} is missing {sorted(missing)[0]!r}.")
    if unknown:
        _fail("semantic_bibliography.shape_invalid", f"Bibliography {location} has unknown key {sorted(unknown)[0]!r}.")


def _fail(code: str, message: str) -> NoReturn:
    raise SemanticBibliographyResourceError(code=code, message=message)


__all__ = [
    "SEMANTIC_BIBLIOGRAPHY_MAX_BYTES",
    "SEMANTIC_BIBLIOGRAPHY_MEDIA_TYPE",
    "SEMANTIC_BIBLIOGRAPHY_SCHEMA",
    "SemanticBibliographyResourceError",
    "parse_semantic_bibliography",
]
