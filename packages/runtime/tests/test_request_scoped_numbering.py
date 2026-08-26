"""Request snapshot ownership for numbering schemes and heading cleanup."""

from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from threading import Barrier, Lock
from typing import Any

import pytest

from docwen_core.models.manifest import PluginManifest, RouteSpec
from docwen_core.models.result import ConversionResult

pytestmark = pytest.mark.contract


def _numbering_snapshot(scheme_id: str, prefix: str) -> dict[str, Any]:
    return {
        "gui": {"language": {"locale": "zh_CN"}},
        "numbering": {
            "add": {
                "settings": {"order": [scheme_id]},
                "schemes": {
                    scheme_id: {
                        "name": scheme_id,
                        "description": "request snapshot probe",
                        "enabled": True,
                        "is_system": False,
                        "locales": ["*"],
                        "level_1": {"format": f"{prefix}-{{1.arabic_half}} "},
                    }
                },
            },
            "cleanup": {
                "settings": {"order": [f"{scheme_id}_cleanup"]},
                "rules": [
                    {
                        "id": f"{scheme_id}_cleanup",
                        "enabled": True,
                        "pattern": rf"^{prefix}:\s*",
                        "level": 1,
                    }
                ],
            },
        },
    }


def _write_minimal_config_root(root: Path, marker: str) -> tuple[Path, Path]:
    from docwen_runtime.config.registry import CONFIG_FILES

    configs_dir = root / "configs"
    user_dir = root / "user"
    for spec in CONFIG_FILES:
        path = configs_dir / spec.rel_path
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text("\n", encoding="utf-8")
    (configs_dir / "numbering" / "add.toml").write_text(
        '[settings]\norder = ["shared"]\n'
        "[schemes.shared]\n"
        'name = "shared"\n'
        'description = "loader-root probe"\n'
        "enabled = true\n"
        "is_system = false\n"
        'locales = ["*"]\n'
        "[schemes.shared.level_1]\n"
        f'format = "{marker}-{{1.arabic_half}} "\n',
        encoding="utf-8",
    )
    (configs_dir / "numbering" / "cleanup.toml").write_text(
        '[settings]\norder = ["shared_cleanup"]\n'
        "[[rules]]\n"
        'id = "shared_cleanup"\n'
        "enabled = true\n"
        f"pattern = '^{marker}:\\s*'\n"
        "level = 1\n",
        encoding="utf-8",
    )
    return configs_dir, user_dir


class _CaptureNumberingPlugin:
    def __init__(self, *, barrier: Barrier | None = None) -> None:
        self.seen: dict[
            str,
            tuple[str, str, tuple[str, str], tuple[str, str], bool],
        ] = {}
        self._barrier = barrier
        self._lock = Lock()
        self._manifest = PluginManifest(
            plugin_id="request_numbering_probe",
            name="Request numbering probe",
            version="1.0",
            description="captures request-scoped numbering state",
            routes=[
                RouteSpec(
                    source_format="markdown",
                    target_format="md",
                    action_name="request_numbering_probe",
                    label="request numbering probe",
                )
            ],
        )

    @property
    def manifest(self) -> PluginManifest:
        return self._manifest

    def can_handle(self, source_format: str, target_format: str, action_name: str = "") -> bool:
        return source_format == "markdown" and target_format == "md" and action_name == "request_numbering_probe"

    def convert(self, context: Any) -> ConversionResult:
        from docwen_core.text.heading_numbering import strip_heading_prefix

        if self._barrier is not None:
            self._barrier.wait(timeout=5)
        schemes = context.numbering_registry.list_schemes()
        scheme = schemes[0] if schemes else None
        scheme_id = scheme.scheme_id if scheme is not None else ""
        scheme_level = scheme.levels["level_1"] if scheme is not None else ""
        probe_prefix = str(context.request.options.get("probe_prefix", scheme_id))
        other_prefix = str(context.request.options.get("other_prefix", "UNMATCHED"))
        own_result = strip_heading_prefix(
            f"{probe_prefix}: Title",
            rules=context.heading_cleanup_rules,
        )
        other_result = strip_heading_prefix(
            f"{other_prefix}: Title",
            rules=context.heading_cleanup_rules,
        )
        captured = (
            scheme_id,
            scheme_level,
            own_result,
            other_result,
            isinstance(context.heading_cleanup_rules, tuple),
        )
        with self._lock:
            self.seen[context.request.request_id] = captured
        return ConversionResult(task_id=context.request.request_id, success=False)


def _task_manager(tmp_path: Any, plugin: _CaptureNumberingPlugin) -> Any:
    from docwen_runtime.engine.route_resolver import RouteResolver
    from docwen_runtime.engine.task_manager import TaskManager
    from docwen_runtime.numbering import NumberingSchemeRegistry
    from docwen_runtime.output.finalizer import OutputFinalizer
    from docwen_runtime.plugin_registry.registry import PluginRegistry
    from docwen_runtime.workspace.manager import WorkspaceManager

    plugins = PluginRegistry()
    plugins.register(plugin)
    locale_path = Path(tmp_path) / "locales" / "zh_CN.toml"
    locale_path.parent.mkdir(parents=True, exist_ok=True)
    locale_path.write_text("", encoding="utf-8")
    numbering_registry = NumberingSchemeRegistry.from_config_snapshot(
        {},
        locale_path=locale_path,
    )
    return TaskManager(
        plugins,
        RouteResolver(plugins),
        WorkspaceManager(root_dir=str(tmp_path / "workspace")),
        OutputFinalizer(),
        numbering_registry=numbering_registry,
    )


def test_task_manager_rebuilds_numbering_and_cleanup_from_each_request_snapshot(tmp_path) -> None:
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest

    source = tmp_path / "probe.md"
    source.write_text("# probe\n", encoding="utf-8")
    plugin = _CaptureNumberingPlugin()
    manager = _task_manager(tmp_path, plugin)

    for sequence, (scheme_id, prefix) in enumerate((("A", "A"), ("B", "B"), ("A", "A"))):
        manager.execute_single(
            ConversionRequest(
                request_id=f"request-{sequence}-{scheme_id}",
                input_refs=[FileRef(path=str(source), format="markdown", category="markdown")],
                target_format="md",
                action_name="request_numbering_probe",
                config_snapshot=_numbering_snapshot(scheme_id, prefix),
            )
        )

    assert plugin.seen == {
        "request-0-A": (
            "A",
            "A-{1.arabic_half} ",
            ("A: ", "Title"),
            ("", "UNMATCHED: Title"),
            True,
        ),
        "request-1-B": (
            "B",
            "B-{1.arabic_half} ",
            ("B: ", "Title"),
            ("", "UNMATCHED: Title"),
            True,
        ),
        "request-2-A": (
            "A",
            "A-{1.arabic_half} ",
            ("A: ", "Title"),
            ("", "UNMATCHED: Title"),
            True,
        ),
    }


def test_cleanup_rule_builder_orders_rules_and_treats_empty_snapshot_as_explicit_empty() -> None:
    from docwen_runtime.config import build_heading_cleanup_rules

    snapshot = _numbering_snapshot("request", "REQ")
    snapshot["numbering"]["cleanup"] = {
        "settings": {"order": ["second", "first"]},
        "rules": [
            {"id": "first", "enabled": True, "pattern": "^ONE", "level": 1},
            {"id": "second", "enabled": True, "pattern": "^TWO", "level": 2},
        ],
    }

    rules = build_heading_cleanup_rules(snapshot)

    assert [rule[0] for rule in rules] == ["second", "first"]
    assert build_heading_cleanup_rules({"numbering": {"cleanup": {"rules": []}}}) == ()


def test_concurrent_requests_keep_numbering_and_cleanup_snapshots_isolated(tmp_path) -> None:
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest

    source = tmp_path / "concurrent.md"
    source.write_text("# concurrent\n", encoding="utf-8")
    plugin = _CaptureNumberingPlugin(barrier=Barrier(2))
    manager = _task_manager(tmp_path, plugin)
    requests = [
        ConversionRequest(
            request_id=f"concurrent-{marker}",
            input_refs=[FileRef(path=str(source), format="markdown", category="markdown")],
            target_format="md",
            action_name="request_numbering_probe",
            options={
                "probe_prefix": marker,
                "other_prefix": "B" if marker == "A" else "A",
            },
            config_snapshot=_numbering_snapshot("shared", marker),
        )
        for marker in ("A", "B")
    ]

    with ThreadPoolExecutor(max_workers=2) as executor:
        results = list(executor.map(manager.execute_single, requests))

    assert all(
        not result.success
        and result.error is not None
        and result.error.error_type == "conversion_failed"
        and result.error.diagnostic_code == "PLUGIN_REPORTED_FAILURE"
        for result in results
    )
    assert plugin.seen == {
        "concurrent-A": (
            "shared",
            "A-{1.arabic_half} ",
            ("A: ", "Title"),
            ("", "B: Title"),
            True,
        ),
        "concurrent-B": (
            "shared",
            "B-{1.arabic_half} ",
            ("B: ", "Title"),
            ("", "A: Title"),
            True,
        ),
    }


def test_two_config_loaders_cannot_rebind_a_frozen_request_snapshot(tmp_path) -> None:
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest
    from docwen_runtime.config.loader import ConfigLoader

    configs_a, user_a = _write_minimal_config_root(tmp_path / "root-a", "A")
    configs_b, user_b = _write_minimal_config_root(tmp_path / "root-b", "B")
    loader_a = ConfigLoader(base_dir=configs_a, user_dir=user_a)
    snapshot_a = loader_a.config.as_dict()
    loader_b = ConfigLoader(base_dir=configs_b, user_dir=user_b)
    snapshot_b = loader_b.config.as_dict()

    source = tmp_path / "two-loaders.md"
    source.write_text("# two loaders\n", encoding="utf-8")
    plugin = _CaptureNumberingPlugin()
    manager = _task_manager(tmp_path, plugin)
    for marker, snapshot in (("A", snapshot_a), ("B", snapshot_b)):
        manager.execute_single(
            ConversionRequest(
                request_id=f"loader-{marker}",
                input_refs=[FileRef(path=str(source), format="markdown", category="markdown")],
                target_format="md",
                action_name="request_numbering_probe",
                options={
                    "probe_prefix": marker,
                    "other_prefix": "B" if marker == "A" else "A",
                },
                config_snapshot=snapshot,
            )
        )

    assert plugin.seen == {
        "loader-A": (
            "shared",
            "A-{1.arabic_half} ",
            ("A: ", "Title"),
            ("", "B: Title"),
            True,
        ),
        "loader-B": (
            "shared",
            "B-{1.arabic_half} ",
            ("B: ", "Title"),
            ("", "A: Title"),
            True,
        ),
    }


def test_config_loader_reload_does_not_publish_process_global_cleanup_rules(tmp_path) -> None:
    from docwen_core.text import heading_numbering
    from docwen_runtime.config.loader import ConfigLoader

    configs_dir, user_dir = _write_minimal_config_root(tmp_path / "empty-root", "STALE")
    loader = ConfigLoader(base_dir=configs_dir, user_dir=user_dir)
    (configs_dir / "numbering" / "cleanup.toml").write_text("\n", encoding="utf-8")
    loader.reload()

    for stale_api in (
        "_INJECTED_RULES",
        "_get_strip_rules",
        "reload_clean_rules",
        "set_clean_rules",
        "set_clean_rules_from_data",
    ):
        assert not hasattr(heading_numbering, stale_api)


def test_task_manager_passes_explicit_empty_cleanup_without_fallback(tmp_path) -> None:
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest

    snapshot = _numbering_snapshot("empty", "EMPTY")
    snapshot["numbering"]["cleanup"] = {}
    source = tmp_path / "explicit-empty.md"
    source.write_text("# explicit empty\n", encoding="utf-8")
    plugin = _CaptureNumberingPlugin()
    manager = _task_manager(tmp_path, plugin)

    manager.execute_single(
        ConversionRequest(
            request_id="explicit-empty",
            input_refs=[FileRef(path=str(source), format="markdown", category="markdown")],
            target_format="md",
            action_name="request_numbering_probe",
            options={
                "probe_prefix": "STALE",
                "other_prefix": "EMPTY",
            },
            config_snapshot=snapshot,
        )
    )

    assert plugin.seen == {
        "explicit-empty": (
            "empty",
            "EMPTY-{1.arabic_half} ",
            ("", "STALE: Title"),
            ("", "EMPTY: Title"),
            True,
        )
    }


def test_task_manager_without_snapshot_passes_immutable_empty_cleanup_rules(tmp_path) -> None:
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest

    source = tmp_path / "no-snapshot.md"
    source.write_text("# no snapshot\n", encoding="utf-8")
    plugin = _CaptureNumberingPlugin()
    manager = _task_manager(tmp_path, plugin)

    manager.execute_single(
        ConversionRequest(
            request_id="no-snapshot",
            input_refs=[FileRef(path=str(source), format="markdown", category="markdown")],
            target_format="md",
            action_name="request_numbering_probe",
            options={"probe_prefix": "STALE", "other_prefix": "OTHER"},
        )
    )

    assert plugin.seen == {
        "no-snapshot": (
            "",
            "",
            ("", "STALE: Title"),
            ("", "OTHER: Title"),
            True,
        )
    }
