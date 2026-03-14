"""
策略模块初始化文件

定义了策略注册表和相关的注册、获取函数。
采用基于元数据（Source->Target）的注册机制，支持精确匹配和类别匹配。
"""

import logging

from docwen.errors import StrategyNotFoundError
from docwen.formats import (
    CATEGORY_DOCUMENT,
    CATEGORY_IMAGE,
    CATEGORY_LAYOUT,
    CATEGORY_MARKDOWN,
    CATEGORY_SPREADSHEET,
    CATEGORY_UNKNOWN,
    get_strategy_category_from_format,
)
from docwen.services.strategies.registry import load_all, load_for_action, load_for_category

logger = logging.getLogger(__name__)

__all__ = [
    "CATEGORY_DOCUMENT",
    "CATEGORY_IMAGE",
    "CATEGORY_LAYOUT",
    "CATEGORY_MARKDOWN",
    "CATEGORY_SPREADSHEET",
    "CATEGORY_UNKNOWN",
    "get_action_registry",
    "get_conversion_registry",
    "get_registry_snapshot",
    "get_strategy",
    "load_all",
    "load_for_action",
    "load_for_category",
    "register_action",
    "register_conversion",
    "restore_registry_snapshot",
    "set_registries",
]

# ==================== 策略注册表 ====================
# 转换策略注册表: (source_format, target_format) -> StrategyClass
_conversion_registry: dict[tuple[str, str], type] = {}

# 命名动作注册表: action_name -> StrategyClass
_action_registry: dict[str, type] = {}

_original_conversion_registry: dict[tuple[str, str], type] = _conversion_registry
_original_action_registry: dict[str, type] = _action_registry


def get_action_registry() -> dict[str, type]:
    return _action_registry


def get_conversion_registry() -> dict[tuple[str, str], type]:
    return _conversion_registry


def get_registry_snapshot() -> tuple[
    dict[str, type],
    dict[tuple[str, str], type],
    dict[str, type],
    dict[tuple[str, str], type],
]:
    return _action_registry, _conversion_registry, _original_action_registry, _original_conversion_registry


def restore_registry_snapshot(
    snapshot: tuple[
        dict[str, type],
        dict[tuple[str, str], type],
        dict[str, type],
        dict[tuple[str, str], type],
    ],
) -> None:
    global _action_registry, _conversion_registry, _original_action_registry, _original_conversion_registry
    _action_registry, _conversion_registry, _original_action_registry, _original_conversion_registry = snapshot


def set_registries(
    *,
    action_registry: dict[str, type],
    conversion_registry: dict[tuple[str, str], type],
    set_as_original: bool = False,
) -> None:
    global _action_registry, _conversion_registry, _original_action_registry, _original_conversion_registry
    _action_registry = action_registry
    _conversion_registry = conversion_registry
    if set_as_original:
        _original_action_registry = _action_registry
        _original_conversion_registry = _conversion_registry


def register_conversion(source_format: str, target_format: str):
    """
    注册一个特定格式转换的策略。

    参数:
        source_format: 源格式 (如 'docx', 'image', 'document')
        target_format: 目标格式 (如 'pdf', 'image')

    说明:
        支持注册具体格式 (如 'docx') 或 通用类别 (如 CATEGORY_DOCUMENT)。
        查找时优先精确匹配，其次尝试类别匹配。
    """

    def decorator(cls):
        key = (source_format.lower(), target_format.lower())
        _conversion_registry[key] = cls
        # logger.debug(f"注册转换策略: {source_format} -> {target_format} : {cls.__name__}")
        return cls

    return decorator


def register_action(action_name: str):
    """
    注册一个命名动作 (如 'validate', 'split_pdf').

    参数:
        action_name: 动作标识符
    """

    def decorator(cls):
        _action_registry[action_name] = cls
        # logger.debug(f"注册动作策略: {action_name} : {cls.__name__}")
        return cls

    return decorator


# ==================== 辅助函数 ====================


def _get_category(fmt: str) -> str | None:
    """根据文件格式获取所属类别"""
    category = get_strategy_category_from_format(fmt)
    if category == CATEGORY_UNKNOWN:
        return None
    return category


def get_strategy(
    action_type: str | None = None, source_format: str | None = None, target_format: str | None = None
) -> type:
    """
    智能查找策略类。

    查找逻辑 (优先级从高到低):
    1. 命名动作匹配 (如果提供了 action_type)
    2. 精确格式匹配 (source_format -> target_format)
    3. 源文件类别通用匹配 (Category -> target_format) (例如: document -> pdf)
    4. 纯类别匹配 (Category -> Category) (例如: image -> image)

    参数:
        action_type: 命名动作 (可选, 如 'validate')
        source_format: 源文件格式 (可选, 如 'docx')
        target_format: 目标文件格式 (可选, 如 'pdf')

    返回:
        StrategyClass: 匹配到的策略类

    异常:
        StrategyNotFoundError: 如果未找到匹配的策略
    """
    normalized_action = action_type.strip() if action_type else None
    src = source_format.lower() if source_format else None
    tgt = target_format.lower() if target_format else None

    def _try_once() -> type | None:
        if normalized_action and normalized_action in _action_registry:
            return _action_registry[normalized_action]

        if not (src and tgt):
            return None

        src_cat = _get_category(src)
        tgt_cat = _get_category(tgt)

        if (src, tgt) in _conversion_registry:
            return _conversion_registry[(src, tgt)]

        if src_cat and (src_cat, tgt) in _conversion_registry:
            return _conversion_registry[(src_cat, tgt)]

        if (
            src_cat == CATEGORY_IMAGE
            and tgt_cat == CATEGORY_IMAGE
            and (CATEGORY_IMAGE, CATEGORY_IMAGE) in _conversion_registry
        ):
            return _conversion_registry[(CATEGORY_IMAGE, CATEGORY_IMAGE)]

        return None

    found = _try_once()
    if found:
        return found

    autoload_enabled = (
        _conversion_registry is _original_conversion_registry and _action_registry is _original_action_registry
    )

    if autoload_enabled and normalized_action:
        load_for_action(normalized_action)
        found = _try_once()
        if found:
            return found

    if autoload_enabled and src and tgt:
        load_for_category(_get_category(src))
        found = _try_once()
        if found:
            return found

    if autoload_enabled:
        load_all()
        found = _try_once()
        if found:
            return found

    # 构建错误信息
    error_msg = "没有找到策略: "
    if normalized_action:
        error_msg += f"action='{normalized_action}' "
    if source_format and target_format:
        error_msg += f"conversion='{source_format}->{target_format}'"

    logger.error(error_msg)
    raise StrategyNotFoundError(action_type=action_type, source_format=source_format, target_format=target_format)
