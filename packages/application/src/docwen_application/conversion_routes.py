"""Application-owned route composition for discovery and execution planning."""

from __future__ import annotations

from dataclasses import dataclass

from docwen_application.preconversion.chain_resolver import resolve_chain
from docwen_application.runtime_capability_catalog import RuntimeCapabilityCatalog, RuntimeRoute
from docwen_core.formats import FORMAT_CATEGORY


@dataclass(frozen=True, slots=True)
class ConversionRoutePlan:
    """Canonical routes needed by one admitted source, in execution order.

    The controller executes the same ``resolve_chain`` through its managed
    preconverter. Discovery retains unavailable steps instead of substituting
    a different route or advertising only the final optimizer's availability.
    """

    source_format: str
    source_category: str
    target_format: str
    action_name: str
    routes: tuple[RuntimeRoute, ...]

    @property
    def available(self) -> bool:
        return all(route.available for route in self.routes)

    @property
    def final_route(self) -> RuntimeRoute:
        return self.routes[-1]


def resolve_conversion_route_plan(
    catalog: RuntimeCapabilityCatalog,
    *,
    source_format: str,
    source_category: str,
    target_format: str,
    action_name: str = "",
) -> ConversionRoutePlan | None:
    """Resolve every application step against canonical runtime facts.

    Missing steps return ``None``; contradictory/ambiguous facts retain the
    catalog's typed discovery failure. A same-format action still needs its
    runtime route even though it needs no format preconversion.
    """

    chain = resolve_chain(source_format, target_format, action_name=action_name)
    targets = chain or [target_format]
    routes: list[RuntimeRoute] = []
    current_format, current_category = source_format, source_category
    for index, target in enumerate(targets):
        route = catalog.resolve_route(
            detected_format=current_format,
            workflow_category=current_category,
            action_name=action_name if index == len(targets) - 1 else "",
            target=target,
        )
        if route is None:
            return None
        routes.append(route)
        current_format = target
        current_category = FORMAT_CATEGORY.get(target, current_category)
    return ConversionRoutePlan(source_format, source_category, target_format, action_name, tuple(routes))
