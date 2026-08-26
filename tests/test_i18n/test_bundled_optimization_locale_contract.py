"""Bundled optimization identity is shared by CLI and GUI across locales."""

from __future__ import annotations

import importlib
import json
from functools import cache
from typing import Any

import pytest

pytestmark = pytest.mark.contract


class _BundledController:
    def __init__(self, projection: dict[str, Any]) -> None:
        self._projection = projection

    def describe_runtime_capabilities(self) -> dict[str, Any]:
        return self._projection


@cache
def _bundled_projection() -> dict[str, Any]:
    from docwen_bundle.runtime_factory import _DEFAULT_PLUGIN_IMPORTS
    from docwen_runtime.capabilities import build_runtime_capability_projection

    manifests = [importlib.import_module(import_path).PLUGIN_MANIFEST for import_path in _DEFAULT_PLUGIN_IMPORTS]
    return build_runtime_capability_projection(
        manifests,
        platform_id="windows",
        egress_guard_status={
            "state": "enforced",
            "installed": True,
            "active": True,
            "scope": "docwen_python_process",
            "policy": "deny_dns_and_ip",
            "mechanism": "cpython_audit_hook",
            "bootstrap": "composition_root",
            "local_transports": ["windows_named_pipe", "unix_domain_socket"],
            "external_processes": "not_managed",
        },
    )


def _optimization_identity(resources: list[dict[str, Any]]) -> tuple[tuple[object, ...], ...]:
    return tuple(
        (
            resource["id"],
            resource["action_name"],
            tuple(
                (
                    binding["scope"],
                    binding["route_id"],
                    binding["source"],
                    binding["source_category"],
                    binding["target"],
                )
                for binding in resource["bindings"]
            ),
        )
        for resource in resources
    )


def test_bundled_cli_optimization_inventory_is_locale_invariant(
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docwen_cli.main import main

    controller = _BundledController(_bundled_projection())
    identities: dict[str, tuple[tuple[object, ...], ...]] = {}
    for locale in ("en_US", "zh_CN"):
        assert (
            main(
                ["resources", "list", "optimizations", "--lang", locale, "--json", "--quiet"],
                controller=controller,
            )
            == 0
        )
        payload = json.loads(capsys.readouterr().out)
        identities[locale] = _optimization_identity(payload["data"]["resources"])

    assert identities["en_US"] == identities["zh_CN"]
    assert [item[0] for item in identities["en_US"]] == ["gongwen", "invoice_cn"]
    assert sum(len(item[2]) for item in identities["en_US"]) == 4


def test_bundled_gui_choices_localize_labels_without_changing_identity() -> None:
    from docwen_gui.i18n import get_locale, set_locale, t
    from docwen_gui.view_models._optimization_filter import discover_optimization_choices

    controller = _BundledController(_bundled_projection())
    previous = get_locale()
    identities: dict[str, tuple[tuple[object, ...], ...]] = {}
    labels: dict[str, dict[str, str]] = {}
    try:
        for locale in ("en_US", "zh_CN"):
            set_locale(locale)
            result = discover_optimization_choices(controller, locale=locale)
            assert result.status == "ready"
            identities[locale] = tuple(
                (
                    choice.id,
                    choice.action_name,
                    tuple(
                        (binding.scope, binding.route_id, binding.source, binding.source_category, binding.target)
                        for binding in choice.bindings
                    ),
                )
                for choice in result.choices
            )
            labels[locale] = {choice.id: choice.label for choice in result.choices}
            assert labels[locale] == {
                resource_id: t(f"cli.interactive.optimization_types.{resource_id}", "")
                for resource_id in ("gongwen", "invoice_cn")
            }
    finally:
        set_locale(previous)

    assert identities["en_US"] == identities["zh_CN"]
    assert labels["en_US"] != labels["zh_CN"]
