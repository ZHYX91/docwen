"""
文档优化类型配置模块

对应配置文件：
    - optimization_config.toml: 优化类型定义

包含：
    - DEFAULT_OPTIMIZATION_CONFIG: 默认优化类型配置
    - OptimizationConfigMixin: 优化类型配置获取方法
"""

from typing import Any

from ..safe_logger import safe_log

# ==============================================================================
#                              默认配置
# ==============================================================================

DEFAULT_OPTIMIZATION_CONFIG = {
    "optimization_config": {
        "settings": {"default_type": "gongwen", "order": ["gongwen"]},
        "types": {
            "gongwen": {
                "name": "公文",
                "description": "针对中国公文格式进行优化处理",
                "locales": ["zh_CN"],
                "enabled": True,
                "is_system": True,
            }
        },
    }
}

# ==============================================================================
#                              Mixin 类
# ==============================================================================


class OptimizationConfigMixin:
    """
    文档优化类型配置获取方法 Mixin

    提供优化类型相关配置的访问方法。
    """

    _configs: dict[str, dict[str, Any]]

    # --------------------------------------------------------------------------
    # 第一层：配置块
    # --------------------------------------------------------------------------

    def get_optimization_config_block(self) -> dict[str, Any]:
        """
        获取整个优化类型配置块

        返回:
            Dict[str, Any]: 优化类型配置字典
        """
        config = self._configs.get("optimization_config", {})
        if not config:
            config = DEFAULT_OPTIMIZATION_CONFIG.get("optimization_config", {})
        return config

    # --------------------------------------------------------------------------
    # 第二层：子表
    # --------------------------------------------------------------------------

    def get_optimization_settings(self) -> dict[str, Any]:
        """
        获取优化类型设置子表

        返回:
            Dict[str, Any]: 优化类型设置字典
        """
        return self.get_optimization_config_block().get("settings", {})

    def get_optimization_types(self) -> dict[str, dict]:
        """
        获取所有优化类型

        返回:
            Dict[str, Dict]: 类型字典，键为类型ID，值为类型配置
        """
        config = self.get_optimization_config_block()
        types = config.get("types", {})
        safe_log.debug("获取优化类型列表，共 %d 个类型", len(types))
        return types

    # --------------------------------------------------------------------------
    # 第三层：具体配置值
    # --------------------------------------------------------------------------

    def get_optimization_type_names(self, locale: str | None = None, scope: str | None = None) -> list[str]:
        """
        获取当前语言下可用的优化类型名称列表（供GUI下拉框使用）

        参数:
            locale: 语言代码，如 "zh_CN"、"en_US"，默认使用当前语言
            scope: 作用域过滤，如 "document_to_md"、"layout_to_md"，不传则不过滤

        返回:
            List[str]: 类型显示名称列表（已过滤不适用的类型）
        """
        from docwen.i18n import get_current_locale, t

        if locale is None:
            locale = get_current_locale()

        types = self.get_optimization_types()
        # 获取类型顺序
        settings = self.get_optimization_settings()
        order = settings.get("order", list(types.keys()))

        type_names = []

        # 按顺序遍历类型
        for type_id in order:
            type_config = types.get(type_id)
            if not isinstance(type_config, dict):
                continue

            # 检查是否启用
            if not type_config.get("enabled", True):
                continue

            # 检查 locales 过滤
            locales = type_config.get("locales", ["*"])
            if "*" not in locales and locale not in locales:
                safe_log.debug("跳过优化类型 %s (locales=%s, 当前=%s)", type_id, locales, locale)
                continue

            # 检查 scopes 过滤
            if scope is not None:
                scopes = type_config.get("scopes", ["*"])
                if "*" not in scopes and scope not in scopes:
                    continue

            # 获取名称（支持 name_key 国际化）
            if "name_key" in type_config:
                name_key = type_config["name_key"]
                name = t(f"action_panel.optimization_types.{name_key}")
                # 如果翻译失败，回退到 name 或 ID
                if name.startswith("[") and name.endswith("]"):
                    name = type_config.get("name", type_id)
            else:
                name = type_config.get("name", type_id)

            type_names.append(name)

        safe_log.debug("获取优化类型名称列表 (locale=%s): %s", locale, type_names)
        return type_names

    def get_localized_optimization_types(self, locale: str | None = None, scope: str | None = None) -> dict[str, str]:
        """
        获取当前语言下可用的优化类型字典（供GUI使用）

        参数:
            locale: 语言代码，如 "zh_CN"、"en_US"，默认使用当前语言
            scope: 作用域过滤，如 "document_to_md"、"layout_to_md"，不传则不过滤

        返回:
            Dict[str, str]: 类型字典，键为类型ID，值为显示名称
        """
        from docwen.i18n import get_current_locale, t

        if locale is None:
            locale = get_current_locale()

        types = self.get_optimization_types()
        settings = self.get_optimization_settings()
        order = settings.get("order", list(types.keys()))

        result = {}

        for type_id in order:
            type_config = types.get(type_id)
            if not isinstance(type_config, dict):
                continue

            # 检查是否启用
            if not type_config.get("enabled", True):
                continue

            # 检查 locales 过滤
            locales = type_config.get("locales", ["*"])
            if "*" not in locales and locale not in locales:
                continue

            # 检查 scopes 过滤
            if scope is not None:
                scopes = type_config.get("scopes", ["*"])
                if "*" not in scopes and scope not in scopes:
                    continue

            # 获取名称
            if "name_key" in type_config:
                name_key = type_config["name_key"]
                name = t(f"action_panel.optimization_types.{name_key}")
                if name.startswith("[") and name.endswith("]"):
                    name = type_config.get("name", type_id)
            else:
                name = type_config.get("name", type_id)

            result[type_id] = name

        return result

    def get_default_optimization_type(self) -> str:
        """
        获取默认优化类型ID

        返回:
            str: 默认类型ID
        """
        settings = self.get_optimization_settings()
        type_id = settings.get("default_type", "gongwen")
        safe_log.debug("获取默认优化类型: %s", type_id)
        return type_id

    def get_optimization_type_by_id(self, type_id: str) -> dict[str, Any]:
        """
        根据类型ID获取类型配置

        参数:
            type_id: 类型ID

        返回:
            Dict[str, Any]: 类型配置字典，未找到返回空字典
        """
        types = self.get_optimization_types()
        type_config = types.get(type_id, {})
        if not type_config:
            safe_log.warning("未找到优化类型: %s", type_id)
        return type_config

    def get_optimization_type_by_name(self, type_name: str) -> dict[str, Any]:
        """
        根据类型名称获取类型配置

        参数:
            type_name: 类型显示名称

        返回:
            Dict[str, Any]: 类型配置字典，未找到返回空字典
        """
        types = self.get_optimization_types()
        for _type_id, type_config in types.items():
            if type_config.get("name") == type_name:
                return type_config
        safe_log.warning("未找到名称为 '%s' 的优化类型", type_name)
        return {}

    def has_optimization_types(self, locale: str | None = None, scope: str | None = None) -> bool:
        """
        检查当前语言是否有可用的优化类型

        用于决定是否显示优化类型选项区域。

        参数:
            locale: 语言代码，默认使用当前语言
            scope: 作用域过滤，如 "document_to_md"、"layout_to_md"，不传则不过滤

        返回:
            bool: 是否有可用的优化类型
        """
        type_names = self.get_optimization_type_names(locale, scope)
        return len(type_names) > 0
