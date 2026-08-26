"""Config-driven YAML field processors for MD to DOCX conversion."""

from __future__ import annotations

import importlib
import logging
from collections.abc import Mapping
from typing import Any

logger = logging.getLogger(__name__)


def get_available_processors_from_config(
    config: object,
    current_locale: str | None = None,
) -> list[dict[str, Any]]:
    """Return locale-visible field processors from a config snapshot."""
    field_config = _normalize_field_config(config)
    processors = field_config.get("processors", {})
    if not isinstance(processors, Mapping):
        return []

    locale = current_locale or "zh_CN"
    result: list[dict[str, Any]] = []
    for processor_id in _iter_processor_ids_by_order(field_config):
        cfg = processors.get(processor_id, {})
        if not isinstance(cfg, Mapping):
            continue
        locales = _normalize_locales(cfg.get("locales"))
        if not _locale_matches(locales, locale):
            continue
        module_name = cfg.get("module")
        module = module_name if isinstance(module_name, str) else ""
        item: dict[str, Any] = {
            "id": processor_id,
            "enabled": bool(cfg.get("enabled", True)),
            "locales": locales,
            "module": module,
            "name": str(cfg.get("name", "")),
            "name_key": str(cfg.get("name_key", "")),
            "description": str(cfg.get("description", "")),
            "is_system": bool(cfg.get("is_system", False)),
        }
        if not module:
            item["load_error"] = "missing module"
        result.append(item)
    return result


def run_yaml_processors(
    yaml_data: dict[str, Any],
    config: object,
    current_locale: str | None = None,
) -> None:
    """Run enabled YAML processors in config order, mutating ``yaml_data``."""
    if not yaml_data:
        return
    for processor_id, module in _iter_enabled_processor_modules(config, current_locale):
        processor = _load_yaml_processor_from_module(processor_id, module)
        if processor is None:
            continue
        processor(yaml_data)


def collect_placeholder_rules(
    config: object,
    current_locale: str | None = None,
) -> list[Mapping[str, Any]]:
    """Return placeholder rules exposed by enabled field processor modules."""
    rules: list[Mapping[str, Any]] = []
    for _processor_id, module in _iter_enabled_processor_modules(config, current_locale):
        module_rules = getattr(module, "PLACEHOLDER_RULES", None)
        if isinstance(module_rules, Mapping):
            rules.append(module_rules)
    return rules


def collect_special_placeholder_handlers(
    config: object,
    current_locale: str | None = None,
) -> dict[str, Any]:
    """Return special placeholder handlers from enabled field processors."""
    handlers: dict[str, Any] = {}
    for processor_id, module in _iter_enabled_processor_modules(config, current_locale):
        module_handlers = getattr(module, "SPECIAL_PLACEHOLDER_HANDLERS", None)
        if not isinstance(module_handlers, Mapping):
            continue
        for placeholder_name, handler_ref in module_handlers.items():
            if not isinstance(placeholder_name, str) or not placeholder_name:
                continue
            handler = getattr(module, handler_ref, None) if isinstance(handler_ref, str) else handler_ref
            if callable(handler):
                handlers[placeholder_name] = handler
            else:
                logger.warning(
                    "Field processor special placeholder handler is not callable: %s.%s",
                    processor_id,
                    placeholder_name,
                )
    return handlers


def _iter_enabled_processor_modules(
    config: object,
    current_locale: str | None = None,
) -> list[tuple[str, Any]]:
    field_config = _normalize_field_config(config)
    processors = field_config.get("processors", {})
    if not isinstance(processors, Mapping):
        return []

    locale = current_locale or "zh_CN"
    result: list[tuple[str, Any]] = []
    for processor_id in _iter_processor_ids_by_order(field_config):
        cfg = processors.get(processor_id, {})
        if not isinstance(cfg, Mapping):
            continue
        if not bool(cfg.get("enabled", True)):
            continue
        if not _locale_matches(_normalize_locales(cfg.get("locales")), locale):
            continue
        module_name = cfg.get("module")
        if not isinstance(module_name, str) or not module_name:
            continue
        module = _load_processor_module(processor_id, module_name)
        if module is None:
            continue
        result.append((processor_id, module))
    return result


def _normalize_field_config(config: object) -> Mapping[str, Any]:
    if not isinstance(config, Mapping):
        return {}
    nested = config.get("field_processors")
    if isinstance(nested, Mapping):
        return nested
    return config


def _iter_processor_ids_by_order(config: Mapping[str, Any]) -> list[str]:
    processors = config.get("processors", {})
    if not isinstance(processors, Mapping):
        return []
    settings = config.get("settings", {})
    configured_order = settings.get("order", []) if isinstance(settings, Mapping) else []
    result: list[str] = []
    for item in configured_order:
        if isinstance(item, str) and item in processors and item not in result:
            result.append(item)
    for processor_id in processors:
        if isinstance(processor_id, str) and processor_id not in result:
            result.append(processor_id)
    return result


def _normalize_locales(value: Any) -> list[str]:
    if isinstance(value, list):
        locales = [item for item in value if isinstance(item, str) and item]
        return locales or ["*"]
    if isinstance(value, str) and value:
        return [value]
    return ["*"]


def _locale_matches(locales: list[str], current_locale: str) -> bool:
    return "*" in locales or current_locale in locales


def _load_processor_module(processor_id: str, module_name: str) -> Any | None:
    try:
        return importlib.import_module(module_name)
    except Exception:
        logger.warning(
            "Field processor import failed: %s (%s)",
            processor_id,
            module_name,
            exc_info=True,
        )
        return None


def _load_yaml_processor_from_module(processor_id: str, module: Any) -> Any | None:
    processor = getattr(module, "process_yaml", None)
    if callable(processor):
        return processor

    logger.warning("Field processor has no process_yaml callable: %s (%s)", processor_id, module.__name__)
    return None
