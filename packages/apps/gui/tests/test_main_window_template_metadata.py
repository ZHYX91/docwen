"""Direct tests for MainWindow template metadata normalization."""

from __future__ import annotations

from datetime import datetime

import pytest

from docwen_gui.main_window import _format_template_modified_label

pytestmark = pytest.mark.unit


def test_template_modified_label_formats_registry_nanoseconds() -> None:
    modified_ns = 1_719_914_460_000_000_000

    assert _format_template_modified_label(modified_ns) == datetime.fromtimestamp(modified_ns / 1_000_000_000).strftime(
        "%Y-%m-%d %H:%M"
    )


@pytest.mark.parametrize(
    "modified_ns",
    [None, True, False, 1.5, "1719914460000000000", [], {}, 10**100],
)
def test_template_modified_label_rejects_non_registry_values(modified_ns: object) -> None:
    assert _format_template_modified_label(modified_ns) is None
