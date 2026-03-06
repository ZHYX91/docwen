"""
多语言样式名解析器模块

负责样式名的国际化处理，支持 MD ↔ DOCX 双向转换：
- 注入时：返回当前语言的样式名
- 检测时：返回所有可能的样式名（所有语言 + 第三方别名）

数据源职责划分：
- 语言文件 [styles] 节：我们定义的样式名（用于注入和检测）
- 配置文件 *_aliases：第三方软件内置样式名（仅用于检测）
"""

import logging
from typing import Any, ClassVar

logger = logging.getLogger(__name__)


class StyleNameResolver:
    """
    多语言样式名解析器

    负责样式名的国际化处理，支持 MD ↔ DOCX 双向转换：
    - 注入时：返回当前语言的样式名
    - 检测时：返回所有可能的样式名（所有语言 + 第三方别名）

    使用示例：
        resolver = StyleNameResolver()

        # 注入样式
        name = resolver.get_injection_name("code_block")  # "代码块"

        # 检测样式
        if resolver.is_style_match(para.style.name, "code_block"):
            # 处理代码块
    """

    _instance = None

    # 样式键名到配置文件的映射
    # 用于确定从哪个配置文件加载第三方别名
    STYLE_CONFIG_MAPPING: ClassVar[dict[str, str]] = {
        # 代码样式
        "code_block": "style_code",
        "inline_code": "style_code",
        # 引用样式
        "quote_1": "style_quote",
        "quote_2": "style_quote",
        "quote_3": "style_quote",
        "quote_4": "style_quote",
        "quote_5": "style_quote",
        "quote_6": "style_quote",
        "quote_7": "style_quote",
        "quote_8": "style_quote",
        "quote_9": "style_quote",
        # 公式样式
        "formula_block": "style_formula",
        "inline_formula": "style_formula",
        # 分隔线样式
        "horizontal_rule_1": "style_horizontal_rule",
        "horizontal_rule_2": "style_horizontal_rule",
        "horizontal_rule_3": "style_horizontal_rule",
        # 列表样式
        "list_block": "style_list",
        # 表格样式
        "table_content": "style_table",
        "three_line_table": "style_table",
        "table_grid": "style_table",
    }

    # 样式键名到别名配置键的映射
    # 用于确定从配置文件的哪个字段加载别名
    STYLE_ALIAS_KEY_MAPPING: ClassVar[dict[str, str]] = {
        # 代码样式
        "code_block": "paragraph_style_aliases",
        "inline_code": "character_style_aliases",
        # 引用样式（段落级别）
        "quote_1": "paragraph_style_aliases",
        "quote_2": "paragraph_style_aliases",
        "quote_3": "paragraph_style_aliases",
        "quote_4": "paragraph_style_aliases",
        "quote_5": "paragraph_style_aliases",
        "quote_6": "paragraph_style_aliases",
        "quote_7": "paragraph_style_aliases",
        "quote_8": "paragraph_style_aliases",
        "quote_9": "paragraph_style_aliases",
        # 公式样式
        "formula_block": "paragraph_style_aliases",
        "inline_formula": "character_style_aliases",
        # 表格样式
        "table_content": "table_content_paragraph_style_aliases",
        "three_line_table": "three_line_table_style_aliases",
        "table_grid": "table_grid_style_aliases",
    }

    def __new__(cls):
        """单例模式"""
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        """初始化样式名解析器"""
        if self._initialized:
            return

        self._i18n = None
        self._config_manager = None
        self._alias_cache: dict[str, list[str]] = {}
        self._locale_names_cache: dict[str, list[str]] = {}
        self._detection_names_cache: dict[str, list[str]] = {}

        self._initialized = True
        logger.info("StyleNameResolver 初始化完成")

    @property
    def i18n(self):
        """延迟加载 I18nManager"""
        if self._i18n is None:
            from docwen.i18n.i18n_manager import I18nManager

            self._i18n = I18nManager()
        return self._i18n

    @property
    def config_manager(self):
        """延迟加载 ConfigManager"""
        if self._config_manager is None:
            from docwen.config.config_manager import config_manager

            self._config_manager = config_manager
        return self._config_manager

    # =========================================================================
    # MD → DOCX 方向：注入样式
    # =========================================================================

    def get_injection_name(self, style_key: str) -> str:
        """
        获取当前语言的样式名（用于注入和应用）

        根据当前软件语言设置，返回对应语言的样式名。
        例如：中文环境返回"代码块"，英文环境返回"Code Block"。

        参数：
            style_key: 样式键名，如 "code_block"、"quote_1"

        返回：
            str: 当前语言的样式名

        示例：
            >>> resolver.get_injection_name("code_block")
            "代码块"  # 中文环境
        """
        logger.debug("获取注入样式名: style_key=%s", style_key)

        name = self.i18n.t(f"styles.{style_key}")

        # 检查是否是未找到翻译的标记格式 [key]
        if name.startswith("[") and name.endswith("]"):
            logger.warning("样式键 '%s' 在语言文件中不存在，使用键名作为样式名", style_key)
            return style_key

        logger.debug("样式名解析结果: %s -> %s", style_key, name)
        return name

    # =========================================================================
    # DOCX → MD 方向：检测样式
    # =========================================================================

    def get_all_detection_names(self, style_key: str) -> list[str]:
        """
        获取某样式的所有可能名称（用于检测，带缓存）

        合并来源：
        1. 所有语言文件的 [styles] 节
        2. 配置文件的 *_aliases 列表

        参数：
            style_key: 样式键名

        返回：
            List[str]: 所有可能的样式名列表（去重后）
        """
        # 检查缓存
        if style_key in self._detection_names_cache:
            return self._detection_names_cache[style_key]

        names = set()

        # 1. 添加所有语言版本的样式名
        locale_names = self._get_all_locale_names(style_key)
        names.update(locale_names)

        # 2. 添加第三方别名
        aliases = self._get_style_aliases(style_key)
        names.update(aliases)

        result = list(names)

        # 存入缓存
        self._detection_names_cache[style_key] = result
        logger.debug("检测名称列表: %s -> %s", style_key, result)
        return result

    def is_style_match(self, style_name: str | None, style_key: str) -> bool:
        """
        判断给定样式名是否匹配某个样式键

        参数：
            style_name: 文档中的实际样式名，可能为 None
            style_key: 样式键名

        返回：
            bool: True 表示匹配，False 表示不匹配
        """
        if style_name is None:
            return False

        all_names = self.get_all_detection_names(style_key)
        match = style_name in all_names

        if match:
            logger.debug("样式匹配成功: '%s' -> %s", style_name, style_key)

        return match

    def detect_style_key(self, style_name: str | None, candidate_keys: list[str]) -> str | None:
        """
        从候选样式键列表中检测匹配的样式键

        参数：
            style_name: 文档中的实际样式名
            candidate_keys: 候选样式键列表

        返回：
            Optional[str]: 匹配的样式键，如果都不匹配返回 None
        """
        if style_name is None:
            return None

        for key in candidate_keys:
            if self.is_style_match(style_name, key):
                return key

        return None

    # =========================================================================
    # 引用样式专用方法
    # =========================================================================

    def get_quote_level(self, style_name: str | None) -> int:
        """
        获取引用样式的级别

        参数：
            style_name: 样式名

        返回：
            int: 引用级别（1-9），如果不是引用样式返回 0
        """
        if style_name is None:
            return 0

        # 检查是否是分级引用样式（quote_1 到 quote_9）
        for level in range(1, 10):
            style_key = f"quote_{level}"
            if self.is_style_match(style_name, style_key):
                logger.debug("检测到引用级别: %s -> level=%d", style_name, level)
                return level

        # 检查是否是通用引用样式（使用 quote 别名列表）
        # 通用引用样式默认为 1 级
        aliases = self._get_style_aliases("quote_1")
        if style_name in aliases:
            logger.debug("检测到通用引用样式: %s -> level=1", style_name)
            return 1

        return 0

    def get_quote_injection_name(self, level: int) -> str:
        """
        获取指定级别的引用样式名（用于注入）

        参数：
            level: 引用级别（1-9）

        返回：
            str: 引用样式名
        """
        if level < 1:
            level = 1
        elif level > 9:
            level = 9

        return self.get_injection_name(f"quote_{level}")

    # =========================================================================
    # 内部辅助方法
    # =========================================================================

    def _get_all_locale_names(self, style_key: str) -> list[str]:
        """
        获取所有语言版本的样式名（带缓存）

        参数：
            style_key: 样式键名

        返回：
            List[str]: 各语言版本的样式名列表
        """
        # 检查缓存
        if style_key in self._locale_names_cache:
            return self._locale_names_cache[style_key]

        # 从所有语言获取翻译
        translations = self.i18n.t_all_locales(f"styles.{style_key}")
        names = list(translations.values())
        result = [n for n in names if not n.startswith("[")]

        # 存入缓存
        self._locale_names_cache[style_key] = result
        return result

    def _get_style_aliases(self, style_key: str) -> list[str]:
        """
        获取样式的第三方别名列表

        从配置文件中加载第三方软件（Word/WPS）内置的样式名。

        参数：
            style_key: 样式键名

        返回：
            List[str]: 第三方别名列表
        """
        # 检查缓存
        if style_key in self._alias_cache:
            return self._alias_cache[style_key]

        aliases = []

        # 获取配置文件名
        config_name = self.STYLE_CONFIG_MAPPING.get(style_key)
        if not config_name:
            logger.debug("样式键 %s 没有配置别名文件", style_key)
            self._alias_cache[style_key] = aliases
            return aliases

        # 获取别名字段名
        alias_field = self.STYLE_ALIAS_KEY_MAPPING.get(style_key)
        if not alias_field:
            logger.debug("样式键 %s 没有配置别名字段", style_key)
            self._alias_cache[style_key] = aliases
            return aliases

        # 配置名称到 getter 方法的映射
        CONFIG_GETTER_MAPPING = {
            "style_code": "get_style_code_block",
            "style_quote": "get_style_quote_block",
            "style_formula": "get_style_formula_block",
            "style_table": "get_style_table_block",
        }

        # 从配置文件加载别名
        try:
            getter_name = CONFIG_GETTER_MAPPING.get(config_name)
            if getter_name:
                getter = getattr(self.config_manager, getter_name, None)
                if getter:
                    config = getter()
                    if config:
                        docx_to_md = config.get("docx_to_md", {})
                        aliases = docx_to_md.get(alias_field, [])
                        if isinstance(aliases, list):
                            logger.debug("从 %s 加载别名: %s -> %s", config_name, alias_field, aliases)
                        else:
                            aliases = []
        except Exception as e:
            logger.warning("加载样式别名失败: %s | 错误: %s", style_key, str(e))
            aliases = []

        self._alias_cache[style_key] = aliases
        return aliases

    def clear_cache(self):
        """清除所有缓存"""
        self._alias_cache.clear()
        self._locale_names_cache.clear()
        self._detection_names_cache.clear()
        logger.debug("样式缓存已清除")

    # =========================================================================
    # 样式格式获取方法
    # =========================================================================

    def get_style_format(self, style_key: str) -> dict[str, Any] | None:
        """
        获取样式格式配置

        代理方法，从 I18nManager 获取当前语言的样式格式配置。
        用于样式注入时设置字体、字号、缩进等格式属性。

        参数：
            style_key: 样式键名，如 "body_paragraph"、"heading_1"、"heading_3_9"

        返回：
            Optional[Dict[str, Any]]: 格式配置字典，包含：
                - east_asia_font: 东亚字体（如 "仿宋_GB2312"）
                - ascii_font: 西文字体（如 "Times New Roman"）
                - font_size_pt: 字号（磅值，如 16）
                - first_line_indent_chars: 首行缩进（1/100字符，如 200 表示 2 字符）
                - first_line_indent_cm: 首行缩进（厘米，用于俄语等）
                - spacing_after_twip: 段后间距（twip，1pt=20twip）
                - spacing_before_twip: 段前间距（twip）
                - bold: 是否加粗
                - justification: 对齐方式（"both" 或 "left"）
            如果找不到返回 None

        示例：
            format_config = style_resolver.get_style_format("heading_1")
            # 返回 {"east_asia_font": "黑体", "ascii_font": "Times New Roman", ...}
        """
        logger.debug("获取样式格式配置: style_key=%s", style_key)
        return self.i18n.get_style_format(style_key)


# 全局单例实例
style_resolver = StyleNameResolver()
