from __future__ import annotations

import pytest

import docwen.services.strategies as strategies


pytestmark = pytest.mark.unit


class _S1:
    pass


class _S2:
    pass


class _S3:
    pass


def test_get_strategy_action_takes_priority(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(strategies, "_action_registry", {"validate": _S1}, raising=True)
    monkeypatch.setattr(strategies, "_conversion_registry", {("docx", "pdf"): _S2}, raising=True)

    assert strategies.get_strategy(action_type="validate", source_format="docx", target_format="pdf") is _S1


def test_get_strategy_exact_match(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(strategies, "_action_registry", {}, raising=True)
    monkeypatch.setattr(
        strategies,
        "_conversion_registry",
        {
            ("docx", "pdf"): _S1,
            ("document", "pdf"): _S2,
            ("image", "image"): _S3,
        },
        raising=True,
    )

    assert strategies.get_strategy(source_format="DOCX", target_format="PDF") is _S1


def test_get_strategy_category_to_target_match(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(strategies, "_action_registry", {}, raising=True)
    monkeypatch.setattr(strategies, "_conversion_registry", {("document", "pdf"): _S2}, raising=True)

    assert strategies.get_strategy(source_format="docx", target_format="pdf") is _S2


def test_get_strategy_image_to_image_category_match(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(strategies, "_action_registry", {}, raising=True)
    monkeypatch.setattr(strategies, "_conversion_registry", {("image", "image"): _S3}, raising=True)

    assert strategies.get_strategy(source_format="jpg", target_format="png") is _S3


def test_get_strategy_error_message_contains_context(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(strategies, "_action_registry", {}, raising=True)
    monkeypatch.setattr(strategies, "_conversion_registry", {}, raising=True)

    with pytest.raises(ValueError) as excinfo:
        strategies.get_strategy(source_format="docx", target_format="pdf")

    msg = str(excinfo.value)
    assert "conversion='docx->pdf'" in msg

