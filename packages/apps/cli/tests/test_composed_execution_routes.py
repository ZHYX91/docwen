"""CLI resolves the same full document chain used by application discovery."""

from __future__ import annotations

from types import SimpleNamespace

import pytest

from docwen_application.controller import CapabilityUnavailableError
from docwen_application.runtime_capability_catalog import parse_runtime_capability_catalog
from docwen_cli.commands import convert

from .capability_fixtures import bundled_available_runtime_projection

pytestmark = pytest.mark.contract


@pytest.mark.parametrize("source", ["docx", "doc", "wps", "rtf", "odt"])
@pytest.mark.parametrize("action", ["", "gongwen"])
def test_cli_uses_the_docx_final_route_for_legacy_sources(
    monkeypatch: pytest.MonkeyPatch, source: str, action: str
) -> None:
    projection = bundled_available_runtime_projection()
    monkeypatch.setattr(
        convert, "_inspection_for", lambda *_: SimpleNamespace(detected_format=source, workflow_category="document")
    )
    routes = convert._resolve_runtime_routes(
        parse_runtime_capability_catalog(projection), files=["input"], inspections={}, action=action, target_format="md"
    )
    route = routes["input"]
    assert route.source == "docx" and route.target == "md" and route.action_name == action


@pytest.mark.parametrize("source", ["doc", "wps", "rtf", "odt"])
def test_cli_stops_before_execution_when_preconversion_is_unavailable(
    monkeypatch: pytest.MonkeyPatch, source: str
) -> None:
    projection = bundled_available_runtime_projection()
    group = next(item for item in projection["sources"] if item["id"] == source)
    bridge = next(route for route in group["routes"] if route["target"] == "docx" and route["action"] is None)
    bridge.update(available=False, state="unavailable")
    group["available"] = any(route["available"] for route in group["routes"])
    projection["counts"]["available_routes"] -= 1
    projection["counts"]["unavailable_routes"] += 1
    monkeypatch.setattr(
        convert, "_inspection_for", lambda *_: SimpleNamespace(detected_format=source, workflow_category="document")
    )
    with pytest.raises(CapabilityUnavailableError, match="docx"):
        convert._resolve_runtime_routes(
            parse_runtime_capability_catalog(projection),
            files=["input"],
            inspections={},
            action="gongwen",
            target_format="md",
        )
