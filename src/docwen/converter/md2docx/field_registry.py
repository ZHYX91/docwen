"""
字段处理注册中心

支持不同处理器注册以下能力：
1. YAML 预处理函数
2. 占位符规则
3. 特殊占位符集合
4. 特殊占位符处理器
"""

from __future__ import annotations

import importlib
import logging
import os
from pathlib import Path
from typing import Any

from docwen.config.toml_operations import read_toml_file, update_toml_value
from docwen.translation import get_current_locale
from docwen.utils.path_utils import get_project_root

logger = logging.getLogger(__name__)

RuleType = str
RuleGroup = list[str]
PlaceholderRules = dict[RuleType, list[RuleGroup]]
YamlProcessor = Any
SpecialHandler = Any

DEFAULT_RULES: PlaceholderRules = {
    "delete_paragraph_if_empty": [],
    "delete_cell_if_empty": [],
    "delete_row_if_empty": [],
    "delete_table_if_empty": [],
}

_yaml_processors: dict[str, YamlProcessor] = {}
_placeholder_rules: dict[str, PlaceholderRules] = {}
_special_placeholders: dict[str, set[str]] = {}
_special_handlers: dict[str, SpecialHandler] = {}
_loaded_modules: set[str] = set()
_failed_modules: dict[str, str] = {}

_config_cache: tuple[float, dict[str, Any]] | None = None
_active_processor_cache: dict[tuple[float, str], list[str]] = {}


def register_yaml_processor(processor_id: str, processor: YamlProcessor) -> None:
    _yaml_processors[processor_id] = processor


def register_placeholder_rules(processor_id: str, rules: PlaceholderRules) -> None:
    normalized: PlaceholderRules = {key: list(groups) for key, groups in DEFAULT_RULES.items()}
    for key, groups in rules.items():
        normalized[key] = list(groups)
    _placeholder_rules[processor_id] = normalized


def register_special_placeholders(processor_id: str, placeholders: set[str]) -> None:
    _special_placeholders[processor_id] = set(placeholders)


def register_special_handler(placeholder: str, handler: SpecialHandler) -> None:
    _special_handlers[placeholder] = handler


def _get_config_path() -> Path:
    override = os.environ.get("DOCWEN_FIELD_PROCESSORS_CONFIG")
    if override:
        return Path(override)
    return Path(get_project_root()) / "configs" / "field_processors.toml"


def reset_field_registry_cache() -> None:
    global _config_cache
    _config_cache = None
    _active_processor_cache.clear()
    _failed_modules.clear()


def _load_config() -> tuple[float, dict[str, Any]]:
    global _config_cache
    path = _get_config_path()
    try:
        mtime = path.stat().st_mtime
    except OSError:
        mtime = -1.0

    if _config_cache and _config_cache[0] == mtime:
        return _config_cache[0], _config_cache[1]

    config = read_toml_file(path)
    parsed = config if isinstance(config, dict) else {}
    _config_cache = (mtime, parsed)
    _active_processor_cache.clear()
    _failed_modules.clear()
    return mtime, parsed


def _locale_matches(locales: list[str] | None, current_locale: str) -> bool:
    if not locales:
        return True
    return "*" in locales or current_locale in locales


def _get_effective_locale(current_locale: str | None) -> str:
    if current_locale:
        return current_locale
    try:
        return get_current_locale() or "zh_CN"
    except Exception:
        return "zh_CN"


def _iter_processor_ids_by_order(config: dict[str, Any]) -> list[str]:
    processors = config.get("processors", {})
    if not isinstance(processors, dict):
        return []
    settings = config.get("settings", {})
    ordered = settings.get("order", []) if isinstance(settings, dict) else []
    result: list[str] = []
    for item in ordered:
        if isinstance(item, str) and item in processors and item not in result:
            result.append(item)
    for processor_id in processors:
        if processor_id not in result:
            result.append(processor_id)
    return result


def _ensure_processor_modules_loaded(config: dict[str, Any]) -> None:
    processors = config.get("processors", {})
    if not isinstance(processors, dict):
        return
    for processor_id in _iter_processor_ids_by_order(config):
        cfg = processors.get(processor_id, {})
        if not isinstance(cfg, dict):
            continue
        module_name = cfg.get("module")
        if not isinstance(module_name, str) or not module_name:
            continue
        if module_name in _loaded_modules:
            continue
        if module_name in _failed_modules:
            continue
        try:
            importlib.import_module(module_name)
            _loaded_modules.add(module_name)
        except Exception as e:
            _failed_modules[module_name] = str(e)
            logger.warning("字段处理器模块导入失败: %s (%s)", processor_id, module_name, exc_info=True)


def _get_active_processor_ids(current_locale: str | None = None) -> list[str]:
    mtime, config = _load_config()
    locale = _get_effective_locale(current_locale)
    cached = _active_processor_cache.get((mtime, locale))
    if cached is not None:
        return list(cached)

    _ensure_processor_modules_loaded(config)
    processors = config.get("processors", {})
    if not isinstance(processors, dict):
        return []
    active: list[str] = []
    for processor_id in _iter_processor_ids_by_order(config):
        cfg = processors.get(processor_id, {})
        if not isinstance(cfg, dict):
            continue
        enabled = cfg.get("enabled", True)
        locales = cfg.get("locales", ["*"])
        module_name = cfg.get("module")
        if not isinstance(module_name, str) or not module_name:
            continue
        if isinstance(module_name, str) and module_name in _failed_modules:
            continue
        if enabled and _locale_matches(locales if isinstance(locales, list) else ["*"], locale):
            active.append(processor_id)
    _active_processor_cache[(mtime, locale)] = list(active)
    return active


def _stable_unique_groups(groups: list[RuleGroup]) -> list[RuleGroup]:
    seen: set[tuple[str, ...]] = set()
    unique: list[RuleGroup] = []
    for group in groups:
        key = tuple(group)
        if key in seen:
            continue
        seen.add(key)
        unique.append(group)
    return unique


def get_merged_placeholder_rules(current_locale: str | None = None) -> PlaceholderRules:
    merged: PlaceholderRules = {key: [] for key in DEFAULT_RULES}
    for processor_id in _get_active_processor_ids(current_locale):
        rules = _placeholder_rules.get(processor_id, {})
        for rule_type in DEFAULT_RULES:
            groups = rules.get(rule_type, [])
            if groups:
                merged[rule_type].extend(groups)
    for rule_type in DEFAULT_RULES:
        merged[rule_type] = _stable_unique_groups(merged[rule_type])
    return merged


def get_active_special_placeholders(current_locale: str | None = None) -> set[str]:
    placeholders: set[str] = set()
    for processor_id in _get_active_processor_ids(current_locale):
        placeholders.update(_special_placeholders.get(processor_id, set()))
    return placeholders


def get_active_special_handlers(current_locale: str | None = None) -> dict[str, SpecialHandler]:
    active_placeholders = get_active_special_placeholders(current_locale)
    return {ph: handler for ph, handler in _special_handlers.items() if ph in active_placeholders}


def run_yaml_processors(yaml_data: dict, current_locale: str | None = None) -> None:
    for processor_id in _get_active_processor_ids(current_locale):
        processor = _yaml_processors.get(processor_id)
        if processor:
            processor(yaml_data)


def run_special_handlers(doc, yaml_data: dict, current_locale: str | None = None) -> None:
    for placeholder, handler in get_active_special_handlers(current_locale).items():
        try:
            handler(doc, yaml_data)
        except Exception:
            logger.exception("执行特殊占位符处理器失败: %s", placeholder)


def is_processor_enabled(processor_id: str) -> bool:
    _, config = _load_config()
    processors = config.get("processors", {})
    if not isinstance(processors, dict):
        return False
    cfg = processors.get(processor_id, {})
    if not isinstance(cfg, dict):
        return False
    enabled = cfg.get("enabled", True)
    return bool(enabled)


def set_processor_enabled(processor_id: str, enabled: bool) -> bool:
    path = _get_config_path()
    ok = update_toml_value(path, f"processors.{processor_id}", "enabled", bool(enabled), create_missing=False)
    reset_field_registry_cache()
    return ok


def get_available_processors(current_locale: str | None = None) -> list[dict[str, Any]]:
    _, config = _load_config()
    processors = config.get("processors", {})
    if not isinstance(processors, dict):
        return []
    locale = _get_effective_locale(current_locale)
    order = _iter_processor_ids_by_order(config)
    result: list[dict[str, Any]] = []
    for processor_id in order:
        cfg = processors.get(processor_id, {})
        if not isinstance(cfg, dict):
            continue
        locales = cfg.get("locales", ["*"])
        locales_list = locales if isinstance(locales, list) else ["*"]
        if not _locale_matches(locales_list, locale):
            continue
        module_name = cfg.get("module")
        module = module_name if isinstance(module_name, str) else ""
        item: dict[str, Any] = {
            "id": processor_id,
            "enabled": bool(cfg.get("enabled", True)),
            "locales": locales_list,
            "module": module,
            "name": cfg.get("name", ""),
            "name_key": cfg.get("name_key", ""),
            "description": cfg.get("description", ""),
            "is_system": bool(cfg.get("is_system", False)),
        }
        if not module:
            item["load_error"] = "missing module"
        elif module in _failed_modules:
            item["load_error"] = _failed_modules[module]
        result.append(item)
    return result
