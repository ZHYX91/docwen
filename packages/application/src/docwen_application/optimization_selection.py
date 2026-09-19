"""Select optimization resources using one application conversion plan."""

from __future__ import annotations

from dataclasses import dataclass

from docwen_application.conversion_routes import resolve_conversion_route_plan
from docwen_application.optimization_catalog import OptimizationBinding, OptimizationCatalog, OptimizationResource


@dataclass(frozen=True, slots=True)
class OptimizationSource:
    detected_format: str
    source_category: str

    def __post_init__(self) -> None:
        object.__setattr__(self, "detected_format", self.detected_format.strip().lower())
        object.__setattr__(self, "source_category", self.source_category.strip().lower())


@dataclass(frozen=True, slots=True)
class OptimizationSelection:
    resource: OptimizationResource
    bindings: tuple[OptimizationBinding, ...]
    route_options: tuple[str, ...]


def _binding_for_source(
    catalog: OptimizationCatalog,
    resource: OptimizationResource,
    source: OptimizationSource,
    target: str,
) -> OptimizationBinding | None:
    if not source.detected_format:
        # Settings can show a category without admitting a concrete file.
        return next(
            (
                binding
                for binding in resource.bindings
                if binding.available and binding.target == target and binding.source_category == source.source_category
            ),
            None,
        )
    plan = resolve_conversion_route_plan(
        catalog.runtime_catalog,
        source_format=source.detected_format,
        source_category=source.source_category,
        target_format=target,
        action_name=resource.action_name,
    )
    if plan is None or not plan.available:
        return None
    return next((binding for binding in resource.bindings if binding.route_id == plan.final_route.id), None)


def select_optimizations(
    catalog: OptimizationCatalog,
    *,
    sources: tuple[OptimizationSource, ...] = (),
    target: str = "md",
    configured_order: tuple[str, ...] = (),
    disabled: frozenset[str] = frozenset(),
) -> tuple[OptimizationSelection, ...]:
    """Require a usable plan for every input; intersect actual route options.

    User policy can order/disable known resources, but cannot register them.
    An empty source list lists direct resource bindings, for discovery before
    input admission. It does not assert availability for any particular file.
    """

    by_id = {resource.id: resource for resource in catalog.resources}
    ordered_ids = dict.fromkeys((*configured_order, *by_id))
    selections: list[OptimizationSelection] = []
    for resource_id in ordered_ids:
        resource = by_id.get(resource_id)
        if resource is None or resource_id in disabled or not resource.available:
            continue
        if sources:
            matched = tuple(_binding_for_source(catalog, resource, source, target) for source in sources)
            if any(binding is None for binding in matched):
                continue
            bindings = tuple(binding for binding in matched if binding is not None)
        else:
            bindings = tuple(binding for binding in resource.bindings if binding.available and binding.target == target)
        if not bindings:
            continue
        common = set(catalog.options_for_route(bindings[0].route_id))
        for binding in bindings[1:]:
            common.intersection_update(catalog.options_for_route(binding.route_id))
        options = tuple(option for option in catalog.options_for_route(bindings[0].route_id) if option in common)
        selections.append(OptimizationSelection(resource, bindings, options))
    return tuple(selections)
