"""Shared optimization-catalog parsing and route-join contracts."""

from __future__ import annotations

from copy import deepcopy
from typing import Any

import pytest

from docwen_application.controller import CapabilityUnavailableError
from docwen_application.optimization_catalog import (
    inspect_optimization_catalog,
    parse_optimization_catalog,
)
from docwen_application.optimization_selection import select_optimizations

pytestmark = pytest.mark.unit


def _projection(*, include_resource: bool = True) -> dict[str, Any]:
    runtime = {"state": "available", "platform": "windows"}
    route = {
        "id": "optimizer:docx:md:internal_action",
        "operation": "action",
        "source": "docx",
        "target": "md",
        "action": "internal_action",
        "available": True,
        "state": "available",
        "options": ["locale", "yaml_key_labels"],
    }
    resources = (
        [
            {
                "id": "public-resource",
                "name": "Public resource",
                "action_name": "internal_action",
                "scopes": ["document_to_md"],
                "available": True,
                "state": "available",
                "bindings": [
                    {
                        "scope": "document_to_md",
                        "route_id": route["id"],
                        "source": "docx",
                        "source_category": "document",
                        "target": "md",
                        "available": True,
                        "state": "available",
                    }
                ],
            }
        ]
        if include_resource
        else []
    )
    binding_count = 1 if include_resource else 0
    return {
        "resource": "formats",
        "contract": {"id": "docwen.runtime-capabilities", "version": 1},
        "runtime": runtime,
        "security": {"dependency_egress_guard": {}},
        "gates": [],
        "sources": [{"id": "docx", "category": "document", "available": True, "routes": [route]}],
        "counts": {
            "sources": 1,
            "routes": 1,
            "available_routes": 1,
            "unavailable_routes": 0,
            "actions": 1,
        },
        "optimizations": {
            "resource": "optimizations",
            "contract": {"id": "docwen.optimizations", "version": 1},
            "runtime": runtime,
            "resources": resources,
            "counts": {
                "resources": len(resources),
                "available_resources": len(resources),
                "unavailable_resources": 0,
                "bindings": binding_count,
                "available_bindings": binding_count,
                "unavailable_bindings": 0,
            },
        },
    }


def test_catalog_parses_independent_id_action_and_route_options() -> None:
    catalog = parse_optimization_catalog(_projection())

    resource = catalog.get("public-resource")
    assert resource is not None
    assert resource.action_name == "internal_action"
    assert resource.id != resource.action_name
    assert catalog.compatible_bindings("public-resource", source="docx", target="md") == resource.bindings
    assert catalog.compatible_bindings("public-resource", source="odt", target="md") == ()
    assert catalog.options_for_route(resource.bindings[0].route_id) == ("locale", "yaml_key_labels")
    assert catalog.to_dict() == _projection()["optimizations"]


def test_catalog_result_distinguishes_ready_empty_from_failed() -> None:
    ready = inspect_optimization_catalog(_projection(include_resource=False))

    assert ready.status == "ready"
    assert ready.catalog is not None
    assert ready.catalog.resources == ()
    assert ready.error is None

    malformed = _projection(include_resource=False)
    malformed.pop("optimizations")
    failed = inspect_optimization_catalog(malformed)

    assert failed.status == "failed"
    assert failed.catalog is None
    assert isinstance(failed.error, CapabilityUnavailableError)


@pytest.mark.parametrize(
    ("location", "field", "value", "message"),
    [
        ("route", "action", "guessed_from_resource_id", "disagrees with its route"),
        ("binding", "source", "odt", "disagrees with its route"),
        ("route", "options", "not-a-list", "invalid route options"),
    ],
)
def test_catalog_fails_closed_when_binding_and_route_disagree(
    location: str,
    field: str,
    value: object,
    message: str,
) -> None:
    projection = deepcopy(_projection())
    target = (
        projection["sources"][0]["routes"][0]
        if location == "route"
        else projection["optimizations"]["resources"][0]["bindings"][0]
    )
    target[field] = value

    with pytest.raises(CapabilityUnavailableError, match=message):
        parse_optimization_catalog(projection)


def test_catalog_rejects_success_payload_with_inconsistent_counts() -> None:
    projection = _projection()
    projection["optimizations"]["counts"]["bindings"] = 0

    with pytest.raises(CapabilityUnavailableError, match="inconsistent counts"):
        parse_optimization_catalog(projection)


def test_explicit_gui_locales_select_future_resources_without_changing_public_serialization() -> None:
    projection = _projection()
    projection["optimization_applicable_locales"] = {"public-resource": ["en_US", "ja_JP"]}
    catalog = parse_optimization_catalog(projection)
    assert catalog.to_dict() == projection["optimizations"]
    for locale in ("en_US", "ja_JP"):
        assert [choice.resource.id for choice in select_optimizations(catalog, interface_locale=locale)] == [
            "public-resource"
        ]
    assert select_optimizations(catalog, interface_locale="zh_CN") == ()
    assert select_optimizations(catalog, interface_locale="en_GB") == ()


def test_legacy_public_catalog_remains_explicitly_executable_but_has_no_gui_applicability() -> None:
    catalog = parse_optimization_catalog(_projection())
    assert len(select_optimizations(catalog)) == 1
    assert select_optimizations(catalog, interface_locale="zh_CN") == ()


@pytest.mark.parametrize(
    "metadata",
    [
        None,
        [],
        {"ghost": ["en_US"]},
        {"public-resource": "en_US"},
        {"public-resource": ["en-US"]},
        {"public-resource": ["en_US", "en_US"]},
    ],
)
def test_gui_applicability_rejects_malformed_or_unknown_declarations(metadata: object) -> None:
    projection = _projection()
    projection["optimization_applicable_locales"] = metadata
    with pytest.raises(CapabilityUnavailableError):
        parse_optimization_catalog(projection)
