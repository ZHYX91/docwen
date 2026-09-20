"""Application composition preserves every dependency and canonical route choice."""

from __future__ import annotations

from dataclasses import replace
from types import MappingProxyType

import pytest

from docwen_application.controller import CapabilityUnavailableError
from docwen_application.conversion_routes import resolve_conversion_route_plan
from docwen_application.runtime_capability_catalog import RuntimeCapabilityCatalog, RuntimeRoute, RuntimeSource

pytestmark = pytest.mark.unit


def _route(source: str, target: str, action: str = "", *, available: bool = True) -> RuntimeRoute:
    return RuntimeRoute(
        id=f"{source}:{target}:{action or 'convert'}",
        operation="action" if action else "conversion",
        source=source,
        source_category="document",
        target=target,
        action_name=action,
        available=available,
        state="available" if available else "unavailable",
        options=("locale",) if action else (),
    )


def _catalog(*routes: RuntimeRoute) -> RuntimeCapabilityCatalog:
    return RuntimeCapabilityCatalog(
        contract_id="docwen.runtime-capabilities",
        contract_version=1,
        runtime_state="available",
        platform="test",
        sources=tuple(
            RuntimeSource(
                source,
                "document",
                any(route.available for route in routes if route.source == source),
                tuple(route for route in routes if route.source == source),
            )
            for source in dict.fromkeys(route.source for route in routes)
        ),
        routes_by_id=MappingProxyType({route.id: route for route in routes}),
    )


@pytest.mark.parametrize("source", ["doc", "wps", "rtf", "odt"])
@pytest.mark.parametrize("bridge_available", [True, False])
def test_composed_plan_retains_the_bridge_and_final_optimizer(source: str, bridge_available: bool) -> None:
    bridge = _route(source, "docx", available=bridge_available)
    optimizer = _route("docx", "md", "gongwen")
    plan = resolve_conversion_route_plan(
        _catalog(bridge, optimizer),
        source_format=source,
        source_category="document",
        target_format="md",
        action_name="gongwen",
    )
    assert plan is not None
    assert plan.routes == (bridge, optimizer)
    assert plan.final_route == optimizer
    assert plan.available is bridge_available
    assert plan.source_format == source


@pytest.mark.parametrize("missing", ["bridge", "optimizer"])
def test_missing_step_does_not_form_a_plan(missing: str) -> None:
    route = _route("docx", "md", "gongwen") if missing == "bridge" else _route("odt", "docx")
    assert (
        resolve_conversion_route_plan(
            _catalog(route), source_format="odt", source_category="document", target_format="md", action_name="gongwen"
        )
        is None
    )


def test_unavailable_exact_route_shadows_available_category_fallback() -> None:
    exact = _route("docx", "md", "gongwen", available=False)
    fallback = _route("document", "md", "gongwen")
    plan = resolve_conversion_route_plan(
        _catalog(exact, fallback),
        source_format="docx",
        source_category="document",
        target_format="md",
        action_name="gongwen",
    )
    assert plan is not None and plan.routes == (exact,) and not plan.available


def test_same_format_action_still_resolves_its_runtime_route() -> None:
    route = _route("docx", "docx", "validate")
    plan = resolve_conversion_route_plan(
        _catalog(route),
        source_format="docx",
        source_category="document",
        target_format="docx",
        action_name="validate",
    )
    assert plan is not None and plan.routes == (route,)


def test_ambiguous_route_is_a_discovery_failure() -> None:
    route = _route("docx", "md", "gongwen")
    with pytest.raises(CapabilityUnavailableError, match="ambiguous"):
        resolve_conversion_route_plan(
            _catalog(route, replace(route, id="other-provider")),
            source_format="docx",
            source_category="document",
            target_format="md",
            action_name="gongwen",
        )
