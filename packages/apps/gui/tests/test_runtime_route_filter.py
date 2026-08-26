"""GUI consumption tests for the canonical Runtime capability catalog."""

from __future__ import annotations

from copy import deepcopy

import pytest
from tests.support.gui_vm_fakes import optimization_capability_projection

from docwen_gui.view_models._runtime_route_filter import (
    RuntimeRouteSource,
    discover_runtime_route_choices,
    load_runtime_catalog,
)

pytestmark = pytest.mark.unit


class _Controller:
    def __init__(self, projection: object) -> None:
        self.projection = projection

    def describe_runtime_capabilities(self) -> object:
        return self.projection


def test_load_distinguishes_ready_valid_empty_and_failed() -> None:
    ready = load_runtime_catalog(_Controller(optimization_capability_projection()))
    assert ready.status == "ready"

    empty_projection = deepcopy(optimization_capability_projection())
    empty_projection["sources"] = []
    empty_projection["counts"] = {
        "sources": 0,
        "routes": 0,
        "available_routes": 0,
        "unavailable_routes": 0,
        "actions": 0,
    }
    empty = load_runtime_catalog(_Controller(empty_projection))
    assert empty.status == "empty"
    assert empty.error is None

    failed = load_runtime_catalog(_Controller({"sources": []}))
    assert failed.status == "failed"
    assert failed.error is not None


def test_discovers_concrete_route_targets_and_canonical_options() -> None:
    result = discover_runtime_route_choices(
        _Controller(optimization_capability_projection()),
        sources=(RuntimeRouteSource("docx", "document"),),
        operation="conversion",
    )

    assert result.status == "ready"
    assert result.targets[:3] == ("md", "doc", "odt")
    assert result.get("md") is not None
    assert result.get("md").options == ("locale", "yaml_key_labels")  # type: ignore[union-attr]


def test_batch_choices_use_target_and_option_intersection() -> None:
    result = discover_runtime_route_choices(
        _Controller(optimization_capability_projection()),
        sources=(
            RuntimeRouteSource("xlsx", "spreadsheet"),
            RuntimeRouteSource("csv", "spreadsheet"),
        ),
        operation="conversion",
    )

    assert result.status == "ready"
    assert result.targets == ("xls", "ods", "pdf")
    assert all(choice.options == () for choice in result.choices)


def test_action_matching_does_not_guess_target_or_legacy_action_alias() -> None:
    controller = _Controller(optimization_capability_projection())
    valid = discover_runtime_route_choices(
        controller,
        sources=(RuntimeRouteSource("docx", "document"),),
        operation="action",
        action_name="validate",
    )
    alias = discover_runtime_route_choices(
        controller,
        sources=(RuntimeRouteSource("docx", "document"),),
        operation="action",
        action_name="validate_legacy",
    )

    assert valid.targets == ("docx",)
    assert alias.status == "empty"


def test_duplicate_action_target_signature_fails_closed() -> None:
    projection = deepcopy(optimization_capability_projection())
    docx_source = next(source for source in projection["sources"] if source["id"] == "docx")  # type: ignore[index]
    duplicate = deepcopy(next(route for route in docx_source["routes"] if route["action"] == "validate"))
    duplicate["id"] = "docx:docx:validate:duplicate"
    docx_source["routes"].append(duplicate)
    projection["counts"]["routes"] += 1  # type: ignore[index,operator]
    projection["counts"]["available_routes"] += 1  # type: ignore[index,operator]
    projection["counts"]["actions"] += 1  # type: ignore[index,operator]

    result = discover_runtime_route_choices(
        _Controller(projection),
        sources=(RuntimeRouteSource("docx", "document"),),
        operation="action",
        action_name="validate",
    )

    assert result.status == "failed"
    assert result.error is not None
