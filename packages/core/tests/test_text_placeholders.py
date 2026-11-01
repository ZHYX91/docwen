"""Placeholder restoration cannot recurse into generated links or authored text."""

from types import SimpleNamespace

import pytest

from docwen_core.text.placeholders import placeholder_namespace, restore_placeholders

pytestmark = pytest.mark.contract


def test_namespace_checks_a_repeated_random_value_against_authored_text(monkeypatch) -> None:
    import docwen_core.text.placeholders as module

    monkeypatch.setattr(module.uuid, "uuid4", lambda: SimpleNamespace(hex="a" * 32))
    authored = "before TOKEN" + "A" * 32 + "XXX after"
    namespace = placeholder_namespace(authored, prefix="TOKEN")
    assert namespace not in authored


def test_restoration_is_once_and_generated_backslashes_stay_literal() -> None:
    assert restore_placeholders("A B", {"A": "B", "B": r"\1"}) == "B " + r"\1"
