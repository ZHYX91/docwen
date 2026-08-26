"""Canonical all-available runtime-route fixtures for CLI request unit tests."""

from __future__ import annotations

import importlib
from collections import defaultdict
from typing import Any

from docwen_application.runtime_capability_catalog import (
    RuntimeCapabilityCatalog,
    parse_runtime_capability_catalog,
)
from docwen_bundle.runtime_factory import _DEFAULT_PLUGIN_IMPORTS
from docwen_core.formats.categories import ALL_CATEGORIES, FORMAT_CATEGORY


def bundled_available_runtime_projection() -> dict[str, Any]:
    """Project the bundled manifests while forcing machine gates available.

    These request-construction tests exercise route identity and option
    schemas, not current-machine dependency probes. The production projection
    and unavailable-state behavior have separate Runtime tests.
    """

    grouped: dict[str, list[dict[str, object]]] = defaultdict(list)
    categories: dict[str, str] = {}
    route_count = 0
    action_count = 0
    for import_path in _DEFAULT_PLUGIN_IMPORTS:
        manifest = importlib.import_module(import_path).PLUGIN_MANIFEST
        for route in manifest.routes:
            source = route.source_format
            category = FORMAT_CATEGORY.get(source, source if source in ALL_CATEGORIES else "other")
            categories[source] = category
            grouped[source].append(
                {
                    "id": (f"{manifest.plugin_id}:{source}:{route.target_format}:{route.action_name or 'convert'}"),
                    "operation": "action" if route.action_name else "conversion",
                    "source": source,
                    "target": route.target_format,
                    "action": route.action_name or None,
                    "available": True,
                    "state": "available",
                    "options": sorted(str(key) for key in route.options_schema.get("properties", {})),
                }
            )
            route_count += 1
            action_count += bool(route.action_name)

    sources = [
        {
            "id": source,
            "category": categories[source],
            "available": True,
            "routes": sorted(grouped[source], key=lambda route: str(route["id"])),
        }
        for source in sorted(grouped)
    ]
    return {
        "resource": "formats",
        "contract": {"id": "docwen.runtime-capabilities", "version": 1},
        "runtime": {"state": "available", "platform": "windows"},
        "security": {"dependency_egress_guard": {}},
        "gates": [],
        "sources": sources,
        "counts": {
            "sources": len(sources),
            "routes": route_count,
            "available_routes": route_count,
            "unavailable_routes": 0,
            "actions": action_count,
        },
    }


def bundled_available_runtime_catalog() -> RuntimeCapabilityCatalog:
    return parse_runtime_capability_catalog(bundled_available_runtime_projection())


__all__ = ["bundled_available_runtime_catalog", "bundled_available_runtime_projection"]
