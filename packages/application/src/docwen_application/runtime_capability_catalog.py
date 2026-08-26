"""Typed, fail-closed consumer view of canonical runtime routes."""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass
from types import MappingProxyType
from typing import Literal, NoReturn, cast

from docwen_application.controller import CapabilityUnavailableError

RouteOperation = Literal["conversion", "action"]
RouteState = Literal["available", "available_with_limits", "unavailable"]

_VALID_OPERATIONS = {"conversion", "action"}
_VALID_STATES = {"available", "available_with_limits", "unavailable"}


def _fail(message: str) -> NoReturn:
    raise CapabilityUnavailableError(message)


def _text(value: object, *, field: str) -> str:
    if not isinstance(value, str) or not value:
        _fail(f"Runtime capability discovery returned an invalid {field}.")
    return value


def _state(value: object, *, field: str) -> RouteState:
    if value not in _VALID_STATES:
        _fail(f"Runtime capability discovery returned an invalid {field}.")
    return cast(RouteState, value)


@dataclass(frozen=True, slots=True)
class RuntimeRoute:
    id: str
    operation: RouteOperation
    source: str
    source_category: str
    target: str
    action_name: str
    available: bool
    state: RouteState
    options: tuple[str, ...]

    @property
    def signature(self) -> tuple[RouteOperation, str, str]:
        """Return the resolver-visible route identity inside one source."""

        return self.operation, self.action_name, self.target


@dataclass(frozen=True, slots=True)
class RuntimeSource:
    id: str
    category: str
    available: bool
    routes: tuple[RuntimeRoute, ...]


@dataclass(frozen=True, slots=True)
class RuntimeCapabilityCatalog:
    contract_id: str
    contract_version: int
    runtime_state: str
    platform: str
    sources: tuple[RuntimeSource, ...]
    routes_by_id: Mapping[str, RuntimeRoute]

    def source(self, source_id: str) -> RuntimeSource | None:
        return next((source for source in self.sources if source.id == source_id), None)

    def source_ids_for(self, detected_format: str, workflow_category: str) -> tuple[str, ...]:
        """Return concrete-first source IDs participating in route lookup."""

        return tuple(
            source_id
            for source_id in dict.fromkeys((detected_format, workflow_category))
            if source_id and self.source(source_id) is not None
        )

    def effective_routes(self, detected_format: str, workflow_category: str) -> tuple[RuntimeRoute, ...]:
        """Project concrete routes plus non-shadowed category fallbacks."""

        routes: list[RuntimeRoute] = []
        signatures: set[tuple[RouteOperation, str, str]] = set()
        for source_id in self.source_ids_for(detected_format, workflow_category):
            source = self.source(source_id)
            if source is None:
                continue
            for route in source.routes:
                if route.signature in signatures:
                    continue
                signatures.add(route.signature)
                routes.append(route)
        return tuple(routes)

    def matching_routes(
        self,
        *,
        detected_format: str,
        workflow_category: str,
        action_name: str,
        target: str,
    ) -> tuple[RuntimeRoute, ...]:
        """Match exactly as Runtime does: concrete source, then category."""

        operation: RouteOperation = "action" if action_name else "conversion"
        for source_id in dict.fromkeys((detected_format, workflow_category)):
            source = self.source(source_id)
            if source is None:
                continue
            matches = tuple(
                route
                for route in source.routes
                if route.operation == operation and route.action_name == action_name and route.target == target
            )
            if matches:
                return matches
        return ()

    def resolve_route(
        self,
        *,
        detected_format: str,
        workflow_category: str,
        action_name: str,
        target: str,
    ) -> RuntimeRoute | None:
        """Resolve one unambiguous canonical route, preserving unavailable state."""

        matches = self.matching_routes(
            detected_format=detected_format,
            workflow_category=workflow_category,
            action_name=action_name,
            target=target,
        )
        if len(matches) > 1:
            _fail(
                "Runtime capability discovery returned ambiguous routes for "
                f"{detected_format or workflow_category} -> {target} (action={action_name!r})."
            )
        return matches[0] if matches else None

    def resolve_action_route(
        self,
        *,
        detected_format: str,
        workflow_category: str,
        action_name: str,
    ) -> RuntimeRoute | None:
        """Resolve the unique target of one named action from Runtime facts.

        Concrete-source routes take precedence over category fallbacks, just
        like Runtime route resolution.  A named action with more than one
        target in the selected source group is ambiguous and fails closed.
        """

        for source_id in dict.fromkeys((detected_format, workflow_category)):
            source = self.source(source_id)
            if source is None:
                continue
            matches: list[RuntimeRoute] = [
                route for route in source.routes if route.operation == "action" and route.action_name == action_name
            ]
            if not matches:
                continue
            if len(matches) > 1:
                _fail(
                    f"Runtime capability discovery returned ambiguous targets for {source_id} (action={action_name!r})."
                )
            return matches[0]
        return None

    def options_for_route(self, route_id: str) -> tuple[str, ...]:
        route = self.routes_by_id.get(route_id)
        return route.options if route is not None else ()


def _parse_route(raw: object, *, source_id: str, source_category: str) -> RuntimeRoute:
    if not isinstance(raw, dict):
        _fail("Runtime capability discovery returned an invalid route.")
    route_id = _text(raw.get("id"), field="route id")
    operation_value = raw.get("operation")
    if operation_value not in _VALID_OPERATIONS:
        _fail(f"Runtime capability discovery returned an invalid route operation: {route_id}")
    operation = cast(RouteOperation, operation_value)
    if raw.get("source") != source_id:
        _fail(f"Runtime capability discovery returned a misplaced route: {route_id}")
    action_value = raw.get("action")
    if operation == "action":
        action_name = _text(action_value, field="route action")
    elif action_value is None:
        action_name = ""
    else:
        _fail(f"Runtime conversion route unexpectedly declares an action: {route_id}")
    available = raw.get("available")
    if not isinstance(available, bool):
        _fail(f"Runtime capability discovery returned invalid route availability: {route_id}")
    state = _state(raw.get("state"), field="route state")
    if available != (state != "unavailable"):
        _fail(f"Runtime capability discovery returned inconsistent route state: {route_id}")
    raw_options = raw.get("options")
    if (
        not isinstance(raw_options, list)
        or not all(isinstance(option, str) and option for option in raw_options)
        or len(raw_options) != len(set(raw_options))
    ):
        _fail(f"Runtime capability discovery returned invalid route options: {route_id}")
    return RuntimeRoute(
        id=route_id,
        operation=operation,
        source=source_id,
        source_category=source_category,
        target=_text(raw.get("target"), field="route target"),
        action_name=action_name,
        available=available,
        state=state,
        options=tuple(raw_options),
    )


def parse_runtime_capability_catalog(
    capability_projection: Mapping[str, object],
) -> RuntimeCapabilityCatalog:
    """Parse canonical runtime routes without reconstructing capability facts."""

    contract = capability_projection.get("contract")
    runtime = capability_projection.get("runtime")
    security = capability_projection.get("security")
    gates = capability_projection.get("gates")
    raw_sources = capability_projection.get("sources")
    raw_counts = capability_projection.get("counts")
    if (
        contract != {"id": "docwen.runtime-capabilities", "version": 1}
        or not isinstance(runtime, dict)
        or runtime.get("state") != "available"
        or not isinstance(security, dict)
        or not isinstance(security.get("dependency_egress_guard"), dict)
        or not isinstance(gates, list)
        or not isinstance(raw_sources, list)
        or not isinstance(raw_counts, dict)
    ):
        _fail("Runtime capability discovery returned an incomplete projection.")

    source_ids: set[str] = set()
    route_index: dict[str, RuntimeRoute] = {}
    sources: list[RuntimeSource] = []
    for raw_source in raw_sources:
        if not isinstance(raw_source, dict):
            _fail("Runtime capability discovery returned an invalid source route group.")
        source_id = _text(raw_source.get("id"), field="source id")
        if source_id in source_ids:
            _fail(f"Runtime capability discovery returned a duplicate source id: {source_id}")
        source_ids.add(source_id)
        source_category = _text(raw_source.get("category"), field="source category")
        raw_routes = raw_source.get("routes")
        source_available = raw_source.get("available")
        if not isinstance(raw_routes, list) or not isinstance(source_available, bool):
            _fail("Runtime capability discovery returned an invalid source route group.")
        routes = tuple(
            _parse_route(route, source_id=source_id, source_category=source_category) for route in raw_routes
        )
        route_signatures: set[tuple[RouteOperation, str, str]] = set()
        for route in routes:
            if route.signature in route_signatures:
                _fail(
                    "Runtime capability discovery returned a duplicate route signature "
                    f"inside source {source_id}: operation={route.operation}, "
                    f"action={route.action_name!r}, target={route.target}."
                )
            route_signatures.add(route.signature)
        if source_available != any(route.available for route in routes):
            _fail(f"Runtime capability discovery returned inconsistent source state: {source_id}")
        for route in routes:
            if route.id in route_index:
                _fail(f"Runtime capability discovery returned a duplicate route id: {route.id}")
            route_index[route.id] = route
        sources.append(
            RuntimeSource(
                id=source_id,
                category=source_category,
                available=source_available,
                routes=routes,
            )
        )

    routes = tuple(route_index.values())
    expected_counts = {
        "sources": len(sources),
        "routes": len(routes),
        "available_routes": sum(route.available for route in routes),
        "unavailable_routes": sum(not route.available for route in routes),
        "actions": sum(route.operation == "action" for route in routes),
    }
    if raw_counts != expected_counts:
        _fail("Runtime capability discovery returned inconsistent route counts.")
    return RuntimeCapabilityCatalog(
        contract_id="docwen.runtime-capabilities",
        contract_version=1,
        runtime_state="available",
        platform=_text(runtime.get("platform"), field="runtime platform"),
        sources=tuple(sources),
        routes_by_id=MappingProxyType(route_index),
    )


__all__ = [
    "RouteOperation",
    "RouteState",
    "RuntimeCapabilityCatalog",
    "RuntimeRoute",
    "RuntimeSource",
    "parse_runtime_capability_catalog",
]
