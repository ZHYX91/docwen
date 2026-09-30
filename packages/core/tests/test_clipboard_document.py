"""Strict recursive clipboard-document model contracts."""

from __future__ import annotations

import json

import pytest

from docwen_core.models.clipboard_document import (
    CLIPBOARD_DOCUMENT_SCHEMA,
    ClipboardDocumentError,
    ClipboardImageRef,
    ClipboardParagraph,
    ClipboardTable,
    ClipboardText,
    clipboard_document_to_bytes,
    clipboard_table_header_shape,
    load_clipboard_document_bytes,
)


def _payload(blocks, resources=None) -> bytes:
    return json.dumps(
        {"schema": CLIPBOARD_DOCUMENT_SCHEMA, "blocks": blocks, "resources": resources or []},
        ensure_ascii=False,
    ).encode("utf-8")


def test_model_accepts_empty_and_explicit_header_tables_without_value_coercion() -> None:
    document = load_clipboard_document_bytes(
        _payload(
            [
                {
                    "type": "table",
                    "rowCount": 2,
                    "columnCount": 2,
                    "cells": [
                        {
                            "row": 0,
                            "column": 0,
                            "rowSpan": 1,
                            "columnSpan": 1,
                            "header": True,
                            "scope": "col",
                            "blocks": [{"type": "paragraph", "inlines": [{"type": "text", "value": ""}]}],
                        },
                        {
                            "row": 0,
                            "column": 1,
                            "rowSpan": 1,
                            "columnSpan": 1,
                            "header": True,
                            "scope": "col",
                            "blocks": [{"type": "paragraph", "inlines": [{"type": "text", "value": "Value"}]}],
                        },
                        {
                            "row": 1,
                            "column": 0,
                            "rowSpan": 1,
                            "columnSpan": 1,
                            "blocks": [{"type": "paragraph", "inlines": [{"type": "text", "value": "00123"}]}],
                        },
                        {
                            "row": 1,
                            "column": 1,
                            "rowSpan": 1,
                            "columnSpan": 1,
                            "blocks": [{"type": "paragraph", "inlines": [{"type": "text", "value": "$A^2$ | < ^"}]}],
                        },
                    ],
                }
            ]
        )
    )
    table = document.blocks[0]
    assert isinstance(table, ClipboardTable)
    assert clipboard_table_header_shape(table) == (1, 0)
    first_data = table.cells[2].blocks[0]
    assert isinstance(first_data, ClipboardParagraph)
    assert first_data.inlines == (ClipboardText("00123"),)
    assert load_clipboard_document_bytes(clipboard_document_to_bytes(document)) == document


def test_model_preserves_nested_table_and_missing_image_node() -> None:
    document = load_clipboard_document_bytes(
        _payload(
            [
                {
                    "type": "table",
                    "rowCount": 1,
                    "columnCount": 1,
                    "cells": [
                        {
                            "row": 0,
                            "column": 0,
                            "rowSpan": 1,
                            "columnSpan": 1,
                            "blocks": [
                                {"type": "paragraph", "inlines": [{"type": "text", "value": "before"}]},
                                {
                                    "type": "table",
                                    "rowCount": 1,
                                    "columnCount": 1,
                                    "cells": [
                                        {
                                            "row": 0,
                                            "column": 0,
                                            "rowSpan": 1,
                                            "columnSpan": 1,
                                            "blocks": [
                                                {
                                                    "type": "paragraph",
                                                    "inlines": [
                                                        {
                                                            "type": "image",
                                                            "resourceId": None,
                                                            "alt": "cat",
                                                            "missingReason": "clipboard_resource_unavailable",
                                                        }
                                                    ],
                                                }
                                            ],
                                        }
                                    ],
                                },
                                {"type": "paragraph", "inlines": [{"type": "text", "value": "after"}]},
                            ],
                        }
                    ],
                }
            ]
        )
    )
    outer = document.blocks[0]
    assert isinstance(outer, ClipboardTable)
    assert isinstance(outer.cells[0].blocks[1], ClipboardTable)
    nested = outer.cells[0].blocks[1]
    assert isinstance(nested, ClipboardTable)
    image_paragraph = nested.cells[0].blocks[0]
    assert isinstance(image_paragraph, ClipboardParagraph)
    assert isinstance(image_paragraph.inlines[0], ClipboardImageRef)


@pytest.mark.parametrize(
    "mutate",
    [
        lambda data: data.update({"unknown": True}),
        lambda data: data["blocks"][0]["cells"].append(
            {
                "row": 0,
                "column": 0,
                "rowSpan": 1,
                "columnSpan": 1,
                "blocks": [],
            }
        ),
    ],
)
def test_model_rejects_unknown_fields_and_overlapping_geometry(mutate) -> None:
    data = {
        "schema": CLIPBOARD_DOCUMENT_SCHEMA,
        "blocks": [
            {
                "type": "table",
                "rowCount": 1,
                "columnCount": 1,
                "cells": [
                    {
                        "row": 0,
                        "column": 0,
                        "rowSpan": 1,
                        "columnSpan": 1,
                        "blocks": [],
                    }
                ],
            }
        ],
        "resources": [],
    }
    mutate(data)
    with pytest.raises(ClipboardDocumentError):
        load_clipboard_document_bytes(json.dumps(data).encode("utf-8"))
