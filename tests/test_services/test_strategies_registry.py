"""services 单元测试。"""

from __future__ import annotations

import pytest

import docwen.services.strategies as strategies
from docwen.errors import StrategyNotFoundError

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


@pytest.mark.parametrize("source_format", ["xps", "caj"])
def test_get_strategy_layout_category_match_for_xps_caj(monkeypatch: pytest.MonkeyPatch, source_format: str) -> None:
    monkeypatch.setattr(strategies, "_action_registry", {}, raising=True)
    monkeypatch.setattr(strategies, "_conversion_registry", {("layout", "md"): _S1}, raising=True)

    assert strategies.get_strategy(source_format=source_format, target_format="md") is _S1


def test_get_strategy_image_to_image_category_match(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(strategies, "_action_registry", {}, raising=True)
    monkeypatch.setattr(strategies, "_conversion_registry", {("image", "image"): _S3}, raising=True)

    assert strategies.get_strategy(source_format="jpg", target_format="png") is _S3


def test_get_strategy_error_message_contains_context(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(strategies, "_action_registry", {}, raising=True)
    monkeypatch.setattr(strategies, "_conversion_registry", {}, raising=True)

    with pytest.raises(StrategyNotFoundError) as excinfo:
        strategies.get_strategy(source_format="docx", target_format="pdf")

    assert excinfo.value.code == "strategy_not_found"
    assert excinfo.value.details is not None
    assert "conversion='docx->pdf'" in excinfo.value.details
