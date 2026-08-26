"""Canonical runtime-route catalog contracts shared by CLI and GUI consumers."""

from __future__ import annotations

from copy import deepcopy
from typing import Any

import pytest

from docwen_application.controller import CapabilityUnavailableError
from docwen_application.runtime_capability_catalog import parse_runtime_capability_catalog

pytestmark = pytest.mark.unit


def _route(
    route_id: str,
    *,
    source: str,
    target: str,
    action: str | None = None,
    options: list[str] | None = None,
) -> dict[str, Any]:
    return {
        "id": route_id,
        "operation": "action" if action else "conversion",
        "source": source,
        "target": target,
        "action": action,
        "available": True,
        "state": "available",
        "options": options or [],
    }


def _projection() -> dict[str, Any]:
    sources = [
        {
            "id": "xlsx",
            "category": "spreadsheet",
            "available": True,
            "routes": [
                _route("xlsx-to-md", source="xlsx", target="md", options=["locale", "yaml_key_labels"]),
            ],
        },
        {
            "id": "spreadsheet",
            "category": "spreadsheet",
            "available": True,
            "routes": [
                _route(
                    "merge-tables",
                    source="spreadsheet",
                    target="xlsx",
                    action="merge_tables",
                    options=["merge_mode"],
                ),
            ],
        },
    ]
    return {
        "resource": "formats",
        "contract": {"id": "docwen.runtime-capabilities", "version": 1},
        "runtime": {"state": "available", "platform": "windows"},
        "security": {"dependency_egress_guard": {}},
        "gates": [],
        "sources": sources,
        "counts": {
            "sources": 2,
            "routes": 2,
            "available_routes": 2,
            "unavailable_routes": 0,
            "actions": 1,
        },
    }


def test_catalog_resolves_concrete_route_before_category_fallback() -> None:
    catalog = parse_runtime_capability_catalog(_projection())

    conversion = catalog.resolve_route(
        detected_format="xlsx",
        workflow_category="spreadsheet",
        action_name="",
        target="md",
    )
    aggregate = catalog.resolve_route(
        detected_format="xlsx",
        workflow_category="spreadsheet",
        action_name="merge_tables",
        target="xlsx",
    )

    assert conversion is not None and conversion.id == "xlsx-to-md"
    assert aggregate is not None and aggregate.id == "merge-tables"
    assert catalog.options_for_route(conversion.id) == ("locale", "yaml_key_labels")
    assert [route.id for route in catalog.effective_routes("xlsx", "spreadsheet")] == [
        "xlsx-to-md",
        "merge-tables",
    ]


def test_catalog_resolves_named_action_target_without_a_consumer_default() -> None:
    catalog = parse_runtime_capability_catalog(_projection())

    route = catalog.resolve_action_route(
        detected_format="xlsx",
        workflow_category="spreadsheet",
        action_name="merge_tables",
    )

    assert route is not None
    assert (route.id, route.target) == ("merge-tables", "xlsx")
    assert (
        catalog.resolve_action_route(
            detected_format="xlsx",
            workflow_category="spreadsheet",
            action_name="missing",
        )
        is None
    )


def test_catalog_rejects_ambiguous_named_action_targets() -> None:
    projection = _projection()
    projection["sources"][1]["routes"].append(
        _route(
            "merge-tables-csv",
            source="spreadsheet",
            target="csv",
            action="merge_tables",
        )
    )
    projection["counts"].update(routes=3, available_routes=3, actions=2)
    catalog = parse_runtime_capability_catalog(projection)

    with pytest.raises(CapabilityUnavailableError, match="ambiguous targets"):
        catalog.resolve_action_route(
            detected_format="xlsx",
            workflow_category="spreadsheet",
            action_name="merge_tables",
        )


def test_catalog_rejects_duplicate_route_signature_inside_one_source() -> None:
    projection = _projection()
    projection["sources"][0]["routes"].append(
        _route(
            "xlsx-to-md-duplicate",
            source="xlsx",
            target="md",
            options=["locale"],
        )
    )
    projection["counts"].update(routes=3, available_routes=3)

    with pytest.raises(CapabilityUnavailableError, match="duplicate route signature"):
        parse_runtime_capability_catalog(projection)


def test_catalog_allows_category_fallback_shadowing_and_keeps_concrete_route() -> None:
    projection = _projection()
    projection["sources"][0]["routes"].append(
        _route(
            "xlsx-merge-tables",
            source="xlsx",
            target="xlsx",
            action="merge_tables",
        )
    )
    projection["counts"].update(routes=3, available_routes=3, actions=2)

    catalog = parse_runtime_capability_catalog(projection)

    assert [route.id for route in catalog.effective_routes("xlsx", "spreadsheet")] == [
        "xlsx-to-md",
        "xlsx-merge-tables",
    ]
    resolved = catalog.resolve_action_route(
        detected_format="xlsx",
        workflow_category="spreadsheet",
        action_name="merge_tables",
    )
    assert resolved is not None and resolved.id == "xlsx-merge-tables"


def test_catalog_accepts_a_successful_empty_composition() -> None:
    projection = _projection()
    projection["sources"] = []
    projection["counts"] = {
        "sources": 0,
        "routes": 0,
        "available_routes": 0,
        "unavailable_routes": 0,
        "actions": 0,
    }

    catalog = parse_runtime_capability_catalog(projection)

    assert catalog.sources == ()
    assert catalog.routes_by_id == {}


@pytest.mark.parametrize(
    ("mutation", "message"),
    [
        (lambda payload: payload["sources"][0]["routes"][0].update(source="csv"), "misplaced route"),
        (lambda payload: payload["counts"].update(routes=0), "inconsistent route counts"),
        (lambda payload: payload["sources"][0]["routes"][0].update(options=["locale", "locale"]), "route options"),
    ],
)
def test_catalog_fails_closed_on_route_projection_drift(mutation: Any, message: str) -> None:
    projection = deepcopy(_projection())
    mutation(projection)

    with pytest.raises(CapabilityUnavailableError, match=message):
        parse_runtime_capability_catalog(projection)
