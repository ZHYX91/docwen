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
from typing import List, Optional, Dict, Any

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
    STYLE_CONFIG_MAPPING = {
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
    }
    
    # 样式键名到别名配置键的映射
    # 用于确定从配置文件的哪个字段加载别名
    STYLE_ALIAS_KEY_MAPPING = {
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
        "table_content": "paragraph_style_aliases",
        "three_line_table": "table_style_aliases",
    }
    
    def __new__(cls):
        """单例模式"""
        if cls._instance is None:
            cls._instance = super(StyleNameResolver, cls).__new__(cls)
            cls._instance._initialized = False
        return cls._instance
    
    def __init__(self):
        """初始化样式名解析器"""
        if self._initialized:
            return
        
        self._i18n = None
        self._config_manager = None
        self._alias_cache: Dict[str, List[str]] = {}
        
        self._initialized = True
        logger.info("StyleNameResolver 初始化完成")
    
    @property
    def i18n(self):
        """延迟加载 I18nManager"""
        if self._i18n is None:
            from gongwen_converter.i18n.i18n_manager import I18nManager
            self._i18n = I18nManager()
        return self._i18n
    
    @property
    def config_manager(self):
        """延迟加载 ConfigManager"""
        if self._config_manager is None:
            from gongwen_converter.config.config_manager import config_manager
            self._config_manager = config_manager
        return self._config_manager
    
    # =========================================================================
    # MD → DOCX 方向：注入样式
    # =========================================================================
    
    def get_injection_name(self, style_key: str) -> str:
        """
        获取当前语言的样式名（用于注入）
        
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
    
    def should_inject(self, doc, style_key: str) -> bool:
        """
        判断是否需要注入该样式
        
        检查模板中是否已存在任何语言版本的样式，
        只有都不存在时才需要注入。
        
        参数：
            doc: Document 对象（python-docx）
            style_key: 样式键名
            
        返回：
            bool: True 表示需要注入，False 表示已存在
        """
        logger.debug("检查是否需要注入样式: style_key=%s", style_key)
        
        # 获取所有语言版本的样式名
        all_names = self._get_all_locale_names(style_key)
        
        # 获取文档中已有的样式名集合
        existing_styles = self._get_document_style_names(doc)
        
        # 检查是否有任何语言版本的样式存在
        for name in all_names:
            if name in existing_styles:
                logger.debug("样式 '%s' 已存在于文档中，无需注入", name)
                return False
        
        logger.debug("样式 '%s' 的所有语言版本都不存在，需要注入", style_key)
        return True
    
    def get_usable_name(self, doc, style_key: str) -> Optional[str]:
        """
        获取文档中可用的样式名（按优先级）
        
        用于应用样式时，优先使用文档中已存在的样式。
        优先级：当前语言 > zh_CN > en_US
        
        参数：
            doc: Document 对象
            style_key: 样式键名
            
        返回：
            Optional[str]: 可用的样式名，如果都不存在返回 None
        """
        logger.debug("获取可用样式名: style_key=%s", style_key)
        
        # 获取文档中已有的样式名集合
        existing_styles = self._get_document_style_names(doc)
        
        # 按优先级检查各语言版本
        priority = self.i18n.get_detection_priority()
        
        for locale in priority:
            name = self.i18n.t_locale(f"styles.{style_key}", locale)
            if not name.startswith("[") and name in existing_styles:
                logger.debug("找到可用样式名: %s (locale=%s)", name, locale)
                return name
        
        logger.debug("未找到可用的样式名: style_key=%s", style_key)
        return None
    
    # =========================================================================
    # DOCX → MD 方向：检测样式
    # =========================================================================
    
    def get_all_detection_names(self, style_key: str) -> List[str]:
        """
        获取某样式的所有可能名称（用于检测）
        
        合并来源：
        1. 所有语言文件的 [styles] 节
        2. 配置文件的 *_aliases 列表
        
        参数：
            style_key: 样式键名
            
        返回：
            List[str]: 所有可能的样式名列表（去重后）
        """
        logger.debug("获取所有检测名称: style_key=%s", style_key)
        
        names = set()
        
        # 1. 添加所有语言版本的样式名
        locale_names = self._get_all_locale_names(style_key)
        names.update(locale_names)
        
        # 2. 添加第三方别名
        aliases = self._get_style_aliases(style_key)
        names.update(aliases)
        
        result = list(names)
        logger.debug("检测名称列表: %s -> %s", style_key, result)
        return result
    
    def is_style_match(self, style_name: Optional[str], style_key: str) -> bool:
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
    
    def detect_style_key(self, style_name: Optional[str], candidate_keys: List[str]) -> Optional[str]:
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
    
    def get_quote_level(self, style_name: Optional[str]) -> int:
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
    
    def _get_all_locale_names(self, style_key: str) -> List[str]:
        """
        获取所有语言版本的样式名
        
        参数：
            style_key: 样式键名
            
        返回：
            List[str]: 各语言版本的样式名列表
        """
        translations = self.i18n.t_all_locales(f"styles.{style_key}")
        names = list(translations.values())
        return [n for n in names if not n.startswith("[")]
    
    def _get_style_aliases(self, style_key: str) -> List[str]:
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
        
        # 从配置文件加载别名
        try:
            config = self.config_manager.get_config(config_name)
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
    
    def _get_document_style_names(self, doc) -> set:
        """
        获取文档中所有样式的名称集合
        
        参数：
            doc: Document 对象
            
        返回：
            set: 样式名集合
        """
        style_names = set()
        
        try:
            for style in doc.styles:
                if style.name:
                    style_names.add(style.name)
        except Exception as e:
            logger.warning("获取文档样式名失败: %s", str(e))
        
        return style_names
    
    def clear_cache(self):
        """清除别名缓存"""
        self._alias_cache.clear()
        logger.debug("样式别名缓存已清除")


# 全局单例实例
style_resolver = StyleNameResolver()
