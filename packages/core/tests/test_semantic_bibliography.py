"""Closed bibliography-resource contract."""

from __future__ import annotations

import json

import pytest

from docwen_core.semantic_bibliography import (
    SEMANTIC_BIBLIOGRAPHY_MAX_BYTES,
    SemanticBibliographyResourceError,
    parse_semantic_bibliography,
)

pytestmark = pytest.mark.contract


def _valid_payload() -> dict[str, object]:
    return {
        "schema": "docwen.semantic_bibliography.v1",
        "entries": [
            {
                "item_id": "smith2025",
                "runs": [
                    {"text": "Smith, A. ", "bold": True},
                    {
                        "text": "Neutral documents",
                        "italic": True,
                        "href": "https://example.org/neutral-documents",
                    },
                ],
            }
        ],
    }


def test_semantic_bibliography_parser_preserves_ordered_rich_runs() -> None:
    fragment = parse_semantic_bibliography(json.dumps(_valid_payload()).encode())

    assert fragment.entries[0].item_id == "smith2025"
    assert [run.text for run in fragment.entries[0].runs] == ["Smith, A. ", "Neutral documents"]
    assert fragment.entries[0].runs[0].bold
    assert fragment.entries[0].runs[1].italic
    assert fragment.entries[0].runs[1].href == "https://example.org/neutral-documents"


@pytest.mark.parametrize(
    ("payload", "code"),
    [
        (b'{"schema":"docwen.semantic_bibliography.v1","schema":"again","entries":[]}', "duplicate_key"),
        (b'{"schema":"docwen.semantic_bibliography.v1","entries":[],"unknown":1}', "shape_invalid"),
        (b'{"schema":"wrong","entries":[]}', "schema_unsupported"),
        (b'{"schema":"docwen.semantic_bibliography.v1","entries":NaN}', "nonfinite_number"),
        (
            b'{"schema":"docwen.semantic_bibliography.v1","entries":[{"item_id":"a","runs":[{"text":"x","href":null}]}]}',
            "shape_invalid",
        ),
        (b"\xff", "invalid_utf8"),
    ],
)
def test_semantic_bibliography_parser_rejects_wire_extensions(payload: bytes, code: str) -> None:
    with pytest.raises(SemanticBibliographyResourceError) as error:
        parse_semantic_bibliography(payload)

    assert error.value.code == f"semantic_bibliography.{code}"


def test_semantic_bibliography_parser_rejects_oversized_resource_before_json() -> None:
    with pytest.raises(SemanticBibliographyResourceError) as error:
        parse_semantic_bibliography(b" " * (SEMANTIC_BIBLIOGRAPHY_MAX_BYTES + 1))

    assert error.value.code == "semantic_bibliography.resource_too_large"


def test_semantic_bibliography_parser_accepts_case_insensitive_http_scheme() -> None:
    payload = _valid_payload()
    entries = payload["entries"]
    assert isinstance(entries, list)
    entry = entries[0]
    assert isinstance(entry, dict)
    runs = entry["runs"]
    assert isinstance(runs, list)
    linked = runs[1]
    assert isinstance(linked, dict)
    linked["href"] = "HTTPS://example.org/neutral-documents"

    fragment = parse_semantic_bibliography(json.dumps(payload).encode())

    assert fragment.entries[0].runs[1].href == "HTTPS://example.org/neutral-documents"
