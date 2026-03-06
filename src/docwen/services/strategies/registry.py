"""
策略注册/加载器

目标：
- 将策略“导入即注册”的副作用从包 import 阶段移到显式/按需加载阶段
- 支持按 action 或 source/target 类别进行局部加载，必要时再全量加载
- 提供轻量校验：不导入模块，仅检查 spec 是否存在
"""

from __future__ import annotations

import importlib
import importlib.util
import sys

from docwen.formats import (
    CATEGORY_DOCUMENT,
    CATEGORY_IMAGE,
    CATEGORY_LAYOUT,
    CATEGORY_MARKDOWN,
    CATEGORY_SPREADSHEET,
)

_CATEGORY_TO_PACKAGE: dict[str, str] = {
    "document": "docwen.services.strategies.document",
    "spreadsheet": "docwen.services.strategies.spreadsheet",
    "layout": "docwen.services.strategies.layout",
    "image": "docwen.services.strategies.image",
    "markdown": "docwen.services.strategies.markdown",
}

_ACTION_TO_MODULE: dict[str, str] = {
    "merge_images_to_tiff": "docwen.services.strategies.image.merge",
    "merge_pdfs": "docwen.services.strategies.layout.operations",
    "split_pdf": "docwen.services.strategies.layout.operations",
    "process_md_numbering": "docwen.services.strategies.operations.md_numbering",
    "validate": "docwen.services.strategies.document.validation",
    "merge_tables": "docwen.services.strategies.operations.merge_tables",
}

_ALL_PACKAGES: set[str] = set(_CATEGORY_TO_PACKAGE.values()) | {
    "docwen.services.strategies.operations",
}

_loaded_modules: set[str] = set()


def validate_registry_specs() -> dict[str, str]:
    """
    校验内置策略清单是否可被解析（不导入模块）。

    Returns:
        Dict[str, str]: key=模块名, value=原因
    """
    failures: dict[str, str] = {}
    allowed_categories = {
        CATEGORY_DOCUMENT,
        CATEGORY_SPREADSHEET,
        CATEGORY_LAYOUT,
        CATEGORY_IMAGE,
        CATEGORY_MARKDOWN,
    }

    for category in sorted(_CATEGORY_TO_PACKAGE.keys()):
        if category not in allowed_categories:
            failures[f"category:{category}"] = "unknown category"

    for action in sorted(_ACTION_TO_MODULE.keys()):
        if not action.strip():
            failures["action:<blank>"] = "empty action name"

    normalized_actions: dict[str, set[str]] = {}
    for action in _ACTION_TO_MODULE:
        normalized = (action or "").strip().lower()
        if not normalized:
            continue
        normalized_actions.setdefault(normalized, set()).add(action)
    for normalized, actions in sorted(normalized_actions.items()):
        if len(actions) > 1:
            failures[f"action:{normalized}"] = f"duplicate normalized action names: {', '.join(sorted(actions))}"

    for action, module_name in _ACTION_TO_MODULE.items():
        normalized_action = (action or "").strip()
        normalized_module = (module_name or "").strip()
        if not normalized_action:
            continue
        if not normalized_module:
            failures[f"action:{normalized_action}"] = "empty module path"

    normalized_categories: dict[str, set[str]] = {}
    for category in _CATEGORY_TO_PACKAGE:
        normalized = (category or "").strip().lower()
        if not normalized:
            continue
        normalized_categories.setdefault(normalized, set()).add(category)
    for normalized, categories in sorted(normalized_categories.items()):
        if len(categories) > 1:
            failures[f"category:{normalized}"] = f"duplicate normalized categories: {', '.join(sorted(categories))}"

    for category, package_name in _CATEGORY_TO_PACKAGE.items():
        normalized_category = (category or "").strip()
        normalized_package = (package_name or "").strip()
        if not normalized_category:
            continue
        if not normalized_package:
            failures[f"category:{normalized_category}"] = "empty package name"

    candidates: set[str] = set(_ALL_PACKAGES) | set(_ACTION_TO_MODULE.values())
    for module_name in sorted(candidates):
        module_name = (module_name or "").strip()
        if not module_name:
            continue
        if importlib.util.find_spec(module_name) is None:
            failures[module_name] = "module spec not found"
    return failures


def validate_registry_runtime() -> dict[str, str]:
    """
    运行期闭环校验：通过重新加载策略清单涉及的模块，检查 action 与类别包确实产生注册副作用。
    """
    failures: dict[str, str] = {}

    from docwen.services import strategies as strategies_pkg

    original_snapshot = strategies_pkg.get_registry_snapshot()
    original_loaded_modules = set(_loaded_modules)

    modules_to_load: set[str] = set(_ALL_PACKAGES) | set(_ACTION_TO_MODULE.values())
    modules_to_load = {m.strip() for m in modules_to_load if m and m.strip()}

    try:
        strategies_pkg.set_registries(action_registry={}, conversion_registry={}, set_as_original=True)
        _loaded_modules.clear()

        already_loaded = [
            name
            for name in sys.modules
            if name.startswith("docwen.services.strategies.") and name not in {"docwen.services.strategies.registry"}
        ]
        for module_name in sorted(already_loaded, key=len, reverse=True):
            try:
                importlib.reload(sys.modules[module_name])
                _loaded_modules.add(module_name)
            except Exception as e:
                failures[module_name] = f"reload failed: {e.__class__.__name__}: {e}"

        for module_name in sorted(modules_to_load):
            try:
                if module_name in sys.modules:
                    importlib.reload(sys.modules[module_name])
                else:
                    importlib.import_module(module_name)
                _loaded_modules.add(module_name)
            except Exception as e:
                failures[module_name] = f"import failed: {e.__class__.__name__}: {e}"

        action_registry = strategies_pkg.get_action_registry()
        for action in sorted(_ACTION_TO_MODULE.keys()):
            normalized = (action or "").strip()
            if not normalized:
                continue
            if action not in action_registry:
                failures[f"action:{action}"] = "action not registered"

        from docwen.formats import CATEGORY_UNKNOWN, get_strategy_category_from_format

        categories_seen: set[str] = set()
        for src, tgt in strategies_pkg.get_conversion_registry():
            src_cat = get_strategy_category_from_format(src)
            tgt_cat = get_strategy_category_from_format(tgt)
            if src_cat != CATEGORY_UNKNOWN:
                categories_seen.add(src_cat)
            if tgt_cat != CATEGORY_UNKNOWN:
                categories_seen.add(tgt_cat)

        required_categories = set(_CATEGORY_TO_PACKAGE.keys())
        for category in sorted(required_categories):
            if category not in categories_seen:
                failures[f"category:{category}"] = "no conversions registered for category"

        return failures
    finally:
        strategies_pkg.restore_registry_snapshot(original_snapshot)
        _loaded_modules.clear()
        _loaded_modules.update(original_loaded_modules)


def load_all() -> None:
    for package_name in sorted(_ALL_PACKAGES):
        _import_once(package_name)
    for module_name in sorted(set(_ACTION_TO_MODULE.values())):
        _import_once(module_name)


def load_for_action(action_type: str) -> None:
    module_name = _ACTION_TO_MODULE.get((action_type or "").strip())
    if not module_name:
        return
    _import_once(module_name)


def load_for_category(category: str | None) -> None:
    if not category:
        return
    package_name = _CATEGORY_TO_PACKAGE.get(category.lower())
    if not package_name:
        return
    _import_once(package_name)


def _import_once(module_name: str) -> None:
    if module_name in _loaded_modules:
        return
    importlib.import_module(module_name)
    _loaded_modules.add(module_name)
