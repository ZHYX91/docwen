"""Runtime factory — assembles the production runtime stack.

This module is the composition root for the DocWen runtime.  It:
    1. Discovers and registers all default plugins.
    2. Creates the RouteResolver, WorkspaceManager, and OutputFinalizer.
    3. Assembles the TaskManager.
    4. Returns a RuntimePortAdapter ready for injection into
       ApplicationController.

This satisfies §5.7 of the rewrite plan: "bundle 负责默认桌面版和 CLI
发行版的组装".
"""

from __future__ import annotations

import logging
import os
import stat
from collections.abc import Callable
from pathlib import Path
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from docwen_core.models.task import TaskEvent
    from docwen_runtime.adapters import RuntimePortAdapter
    from docwen_runtime.config import ConfigLoader
    from docwen_runtime.plugin_registry.registry import PluginRegistry

logger = logging.getLogger(__name__)

_WORKSPACE_ROOT_ENV = "DOCWEN_WORKSPACE_ROOT"
_GOVERNED_WORKSPACE_HEADING = "# DocWen 本地工作区"

# Default plugin import paths — each provides a PluginManifest and a
# Plugin class implementing the ConverterPlugin protocol.
_DEFAULT_PLUGIN_IMPORTS: list[str] = [
    "docwen_plugin_document",
    "docwen_plugin_presentation",
    "docwen_plugin_spreadsheet",
    "docwen_plugin_markup",
    "docwen_plugin_layout",
    "docwen_plugin_print",
    "docwen_plugin_image",
    "docwen_plugin_markdown",
    "docwen_plugin_optimizer_gongwen",
    "docwen_plugin_optimizer_invoice_cn",
    "docwen_plugin_proofread",
]


def create_runtime_port(
    *,
    config_loader: ConfigLoader | None = None,
    extra_plugins: list[str] | None = None,
    event_callback: Callable[[TaskEvent], None] | None = None,
) -> RuntimePortAdapter:
    """Create a fully-wired RuntimePortAdapter with default plugins registered.

    This is the single production entry point for creating a working
    runtime.  Both CLI and GUI should obtain their RuntimePort through
    this factory.

    Args:
        config_loader: Configuration owner selected by the composition root.
            When omitted, this factory creates a fresh loader owned by the
            returned runtime; no process-global loader is consulted.
        extra_plugins: Additional plugin import paths (for custom builds).
        event_callback: Optional task-event callback used by desktop/GUI
            entry points to stream runtime events back to the main thread.

    Returns:
        A RuntimePortAdapter wrapping a fully-configured TaskManager.

    Raises:
        RuntimeError: If a required default plugin cannot be imported or the
            runtime stack otherwise fails to initialise. Additional plugins
            remain optional and are skipped with a warning when unavailable.
    """
    from docwen_runtime.adapters import RuntimePortAdapter
    from docwen_runtime.capabilities import build_runtime_capability_projection
    from docwen_runtime.config import ConfigLoader, build_proofread_rules
    from docwen_runtime.engine.route_resolver import RouteResolver
    from docwen_runtime.engine.task_manager import TaskManager
    from docwen_runtime.numbering.registry import NumberingSchemeRegistry
    from docwen_runtime.output.finalizer import OutputFinalizer
    from docwen_runtime.output.manifest import OutputManifestWriter
    from docwen_runtime.plugin_registry.registry import PluginRegistry
    from docwen_runtime.resources import ResourceRegistry
    from docwen_runtime.workspace.manager import WorkspaceManager

    if config_loader is None:
        config_loader = ConfigLoader()

    # ── 1. Plugin registry ──────────────────────────────────────────
    registry = PluginRegistry()

    required_plugin_imports = list(_DEFAULT_PLUGIN_IMPORTS)
    plugin_imports = list(required_plugin_imports)
    if extra_plugins:
        plugin_imports.extend(extra_plugins)

    loaded_count = 0
    required_failures: list[tuple[str, Exception]] = []
    required_plugin_count = len(required_plugin_imports)
    for index, import_path in enumerate(plugin_imports):
        try:
            _register_plugin(import_path, registry)
            loaded_count += 1
        except Exception as exc:
            logger.warning(
                "Failed to load plugin %s: %s. It will not be available.",
                import_path,
                exc,
            )
            if index < required_plugin_count:
                required_failures.append((import_path, exc))

    if required_failures:
        details = "; ".join(f"{import_path}: {type(exc).__name__}: {exc}" for import_path, exc in required_failures)
        raise RuntimeError(f"Failed to load required default plugins: {details}")

    if loaded_count == 0:
        logger.error("No plugins could be loaded. The runtime will be unable to execute any conversions.")

    # ── 2. Pipeline components ──────────────────────────────────────
    route_resolver = RouteResolver(registry)
    workspace_manager = WorkspaceManager(root_dir=_runtime_workspace_root())
    output_finalizer = OutputFinalizer()

    # ── 2b. Numbering scheme registry ────────────────────────────────
    config_snapshot = config_loader.config.as_dict()
    configured_locale = _configured_locale(config_snapshot)
    locale_path = ResourceRegistry.default().locales_dir() / f"{configured_locale}.toml"
    numbering_registry: Any = NumberingSchemeRegistry.from_config_snapshot(
        config_snapshot,
        locale_path=locale_path,
    )
    proofread_rules: Any = build_proofread_rules(config_snapshot)

    # ── 3. Task manager ─────────────────────────────────────────────
    task_manager = TaskManager(
        plugin_registry=registry,
        route_resolver=route_resolver,
        workspace_manager=workspace_manager,
        output_finalizer=output_finalizer,
        numbering_registry=numbering_registry,
        proofread_rules=proofread_rules,
    )

    # ── 4. Port adapter ─────────────────────────────────────────────
    adapter = RuntimePortAdapter(
        task_manager=task_manager,
        event_callback=event_callback,
        config_loader=config_loader,
        capability_provider=lambda: build_runtime_capability_projection(registry.list_manifests()),
        output_manifest_writer=OutputManifestWriter(output_finalizer),
    )
    logger.info(
        "Runtime port created with %d/%d plugins loaded.",
        loaded_count,
        len(plugin_imports),
    )
    return adapter


def _runtime_workspace_root() -> Path:
    """Select an explicit runtime workspace without using process TEMP.

    Source-tree launches are bound to the engineering workspace. Installed
    distributions use the platform cache directory, which is application-owned
    and never a drive root.
    """
    configured = os.environ.get(_WORKSPACE_ROOT_ENV, "").strip()
    if configured:
        governed_root = Path(configured)
        if not _is_governed_workspace_root(governed_root):
            raise RuntimeError(f"invalid governed DocWen workspace: {governed_root}")
        return governed_root / "temp" / "runtime"

    source_repository = _source_repository(Path(__file__))
    if source_repository is not None:
        governed_root = source_repository.parent.parent / ".workspace"
        if not _is_governed_workspace_root(governed_root):
            raise RuntimeError(f"governed DocWen workspace is missing: {governed_root}")
        return governed_root / "temp" / "runtime"

    from platformdirs import user_cache_dir

    return Path(user_cache_dir("docwen", appauthor=False)) / "runtime"


def _source_repository(module_file: Path) -> Path | None:
    """Return the repository only for the exact editable-source layout."""
    try:
        repository = module_file.resolve().parents[4]
    except IndexError:
        return None
    if repository.name.casefold() != "docwen" or repository.parent.name.casefold() != "repos":
        return None
    return repository


def _is_governed_workspace_root(path: Path) -> bool:
    """Validate the minimum immutable identity used by runtime placement."""
    if not path.is_absolute():
        return False
    try:
        root_metadata = path.lstat()
        temp_metadata = (path / "temp").lstat()
        heading_matches = (path / "README.md").read_text(encoding="utf-8").startswith(_GOVERNED_WORKSPACE_HEADING)
    except OSError:
        return False
    reparse_flag = int(getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0))

    def is_plain_directory(metadata: os.stat_result) -> bool:
        return stat.S_ISDIR(metadata.st_mode) and not (
            stat.S_ISLNK(metadata.st_mode)
            or (reparse_flag and int(getattr(metadata, "st_file_attributes", 0)) & reparse_flag)
        )

    return is_plain_directory(root_metadata) and is_plain_directory(temp_metadata) and heading_matches


def _configured_locale(config_snapshot: dict[str, Any]) -> str:
    gui = config_snapshot.get("gui", {})
    language = gui.get("language", {}) if isinstance(gui, dict) else {}
    locale = language.get("locale") if isinstance(language, dict) else None
    return str(locale).strip() if locale else "zh_CN"


def _register_plugin(import_path: str, registry: PluginRegistry) -> None:
    """Import a plugin package and register its manifest + plugin class.

    Each plugin package must export:
        - ``PLUGIN_MANIFEST``: a ``PluginManifest`` instance.
        - ``PLUGIN_CLASS``: a class implementing ``ConverterPlugin``.

    Missing exports raise ``AttributeError``. The caller decides whether
    that failure is fatal for a required plugin or degradable for an optional
    extra plugin.
    """
    import importlib

    mod = importlib.import_module(import_path)

    manifest = getattr(mod, "PLUGIN_MANIFEST", None)
    plugin_cls = getattr(mod, "PLUGIN_CLASS", None)

    if manifest is None:
        raise AttributeError(f"{import_path} does not export PLUGIN_MANIFEST")
    if plugin_cls is None:
        raise AttributeError(f"{import_path} does not export PLUGIN_CLASS")

    plugin_instance = plugin_cls()
    registry.register(plugin_instance)
    logger.debug("Registered plugin: %s", manifest.plugin_id)


__all__ = ["create_runtime_port"]
