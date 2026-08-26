"""GUI contract tests for Runtime-owned optimization resources."""

from __future__ import annotations

from types import SimpleNamespace
from typing import Any, cast
from unittest.mock import MagicMock

import pytest

from docwen_gui.main_window import _to_markdown_locale_options
from docwen_gui.view_models._optimization_filter import (
    OptimizationSource,
    discover_optimization_choices,
)
from docwen_gui.view_models.action_area_vm import ActionAreaViewModel

pytestmark = pytest.mark.unit


def _binding(source: str, category: str, route_id: str, *, scope: str) -> dict[str, Any]:
    return {
        "scope": scope,
        "route_id": route_id,
        "source": source,
        "source_category": category,
        "target": "md",
        "available": True,
        "state": "available",
    }


def _projection(*, include_resources: bool = True) -> dict[str, Any]:
    routes = [
        ("docx", "document", "docx:md:internal-gongwen", "internal-gongwen", ["locale"]),
        ("pdf", "layout", "pdf:md:invoice-action", "invoice-action", ["locale", "yaml_key_labels"]),
        ("image", "image", "image:md:invoice-action", "invoice-action", ["locale"]),
    ]
    sources = [
        {
            "id": source,
            "category": category,
            "available": True,
            "routes": [
                {
                    "id": route_id,
                    "operation": "action",
                    "action": action,
                    "source": source,
                    "target": "md",
                    "available": True,
                    "state": "available",
                    "options": options,
                }
            ],
        }
        for source, category, route_id, action, options in routes
    ]
    resources = (
        [
            {
                "id": "public-gongwen",
                "name": "Canonical Gongwen",
                "action_name": "internal-gongwen",
                "scopes": ["document_to_md"],
                "available": True,
                "state": "available",
                "bindings": [_binding("docx", "document", routes[0][2], scope="document_to_md")],
            },
            {
                "id": "invoice-resource",
                "name": "Canonical Invoice",
                "action_name": "invoice-action",
                "scopes": ["layout_to_md", "image_to_md"],
                "available": True,
                "state": "available",
                "bindings": [
                    _binding("pdf", "layout", routes[1][2], scope="layout_to_md"),
                    _binding("image", "image", routes[2][2], scope="image_to_md"),
                ],
            },
        ]
        if include_resources
        else []
    )
    binding_count = sum(len(resource["bindings"]) for resource in resources)
    return {
        "resource": "formats",
        "contract": {"id": "docwen.runtime-capabilities", "version": 1},
        "runtime": {"state": "available", "platform": "test"},
        "security": {"dependency_egress_guard": {}},
        "gates": [],
        "sources": sources,
        "counts": {
            "sources": 3,
            "routes": 3,
            "available_routes": 3,
            "unavailable_routes": 0,
            "actions": 3,
        },
        "optimizations": {
            "resource": "optimizations",
            "contract": {"id": "docwen.optimizations", "version": 1},
            "runtime": {"state": "available", "platform": "test"},
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


def _controller(
    *,
    projection: dict[str, Any] | None = None,
    policy: dict[str, Any] | None = None,
) -> MagicMock:
    controller = MagicMock()
    controller.describe_runtime_capabilities.return_value = projection or _projection()
    controller.config_port.get.side_effect = lambda key, default=None: policy if key == "optimize" else default
    return controller


def test_config_orders_and_disables_only_runtime_resources() -> None:
    result = discover_optimization_choices(
        _controller(
            policy={
                "settings": {"order": ["ghost", "invoice-resource", "public-gongwen"]},
                "types": {"public-gongwen": {"enabled": False}, "ghost": {"enabled": True}},
            }
        ),
        locale="en_US",
    )

    assert result.status == "ready"
    assert [choice.id for choice in result.choices] == ["invoice-resource"]


def test_exact_format_and_category_wildcard_binding_are_distinct() -> None:
    controller = _controller()
    docx = discover_optimization_choices(
        controller,
        locale="en_US",
        sources=(OptimizationSource("docx", "document"),),
    )
    odt = discover_optimization_choices(
        controller,
        locale="en_US",
        sources=(OptimizationSource("odt", "document"),),
    )
    png = discover_optimization_choices(
        controller,
        locale="en_US",
        sources=(OptimizationSource("png", "image"),),
    )

    assert [choice.id for choice in docx.choices] == ["public-gongwen"]
    assert odt.choices == ()
    assert [choice.id for choice in png.choices] == ["invoice-resource"]


def test_batch_requires_every_input_and_intersects_route_options() -> None:
    result = discover_optimization_choices(
        _controller(),
        locale="en_US",
        sources=(OptimizationSource("pdf", "layout"), OptimizationSource("png", "image")),
    )

    assert [choice.id for choice in result.choices] == ["invoice-resource"]
    assert result.choices[0].route_options == ("locale",)


def test_ready_empty_and_failed_discovery_remain_distinct() -> None:
    empty = discover_optimization_choices(
        _controller(projection=_projection(include_resources=False)),
        locale="en_US",
    )
    failed = discover_optimization_choices(_controller(projection={"broken": True}), locale="en_US")

    assert empty.status == "ready" and empty.choices == () and empty.error is None
    assert failed.status == "failed" and failed.choices == () and failed.error is not None


def test_action_area_keeps_public_id_separate_from_internal_action() -> None:
    controller = _controller(
        policy={
            "settings": {"order": ["public-gongwen"]},
            "types": {"public-gongwen": {"enabled": True}},
        }
    )
    controller.config_port.get.side_effect = lambda key, default=None: {
        "optimize": {
            "settings": {"order": ["public-gongwen"]},
            "types": {"public-gongwen": {"enabled": True}},
        },
        "document.to_md_enable_optimization": True,
        "document.to_md_optimization_type": "public-gongwen",
    }.get(key, default)
    vm = ActionAreaViewModel(main_vm=cast(Any, SimpleNamespace(controller=controller)))

    vm.setup_for_document_file("/test.docx", "docx")

    assert vm.optimize_for_type == "public-gongwen"
    assert vm.action_name == "internal-gongwen"
    assert vm.optimization_route_options == ("locale",)


def test_route_options_replace_action_name_special_cases(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr("docwen_gui.i18n.get_locale", lambda: "de_DE")

    locale_only = _to_markdown_locale_options(
        {},
        target_format="md",
        action_name="any-action",
        route_options=("locale",),
    )
    with_labels = _to_markdown_locale_options(
        {},
        target_format="md",
        action_name="another-action",
        route_options=("locale", "yaml_key_labels"),
    )
    docx_locale = _to_markdown_locale_options(
        {},
        target_format="docx",
        route_options=("locale",),
    )

    assert locale_only == {"locale": "de_DE"}
    assert with_labels["locale"] == "de_DE"
    assert "yaml_key_labels" in with_labels
    assert docx_locale == {"locale": "de_DE"}
