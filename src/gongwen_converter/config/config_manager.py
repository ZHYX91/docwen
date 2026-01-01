"""
配置管理单例类
提供配置的加载、访问和修改功能，支持GUI配置
"""

import os
from typing import Dict, Any, List, Tuple, Optional, Union
from .safe_logger import safe_log
from .constants import DEFAULT_CONFIG, CONFIG_FILES, DEFAULT_NUMBERING_PATTERNS_CONFIG
from .toml_operations import read_toml_file, update_toml_value

class ConfigManager:
    """
    配置管理单例类（无日志依赖）
    配置加载：直接使用配置文件中的值（即使为空）
    """
    
    _instance = None
    _configs: Dict[str, Dict[str, Any]] = {}
    _initialized = False
    _config_dir = ""
    
    def __new__(cls, config_dir: str = None):
        if cls._instance is None:
            cls._instance = super(ConfigManager, cls).__new__(cls)
            cls._instance._initialize(config_dir or "configs")
        return cls._instance
    
    def _initialize(self, config_dir: str):
        """初始化配置管理器（无日志依赖）"""
        if self._initialized:
            return
            
        # 使用路径工具获取配置目录
        try:
            from gongwen_converter.utils.path_utils import get_config_path, get_project_root
            # 如果传入的是相对路径，则基于项目根目录解析
            if not os.path.isabs(config_dir):
                self._config_dir = os.path.join(get_project_root(), config_dir)
            else:
                self._config_dir = os.path.normpath(config_dir)
        except ImportError:
            # 如果路径工具不可用，回退到旧逻辑
            self._config_dir = os.path.normpath(config_dir)
        
        self._configs = {}
        
        # 修改：即使配置目录不存在，也加载默认配置
        if not os.path.isdir(self._config_dir):
            safe_log.warning("配置目录不存在: %s，使用默认配置", self._config_dir)
        
        # 加载所有配置文件（自动合并默认值）
        self._load_all_configs()
        self._initialized = True
        safe_log.info("配置管理器初始化完成 | 目录: %s", self._config_dir)
    
    def _load_all_configs(self):
        """加载所有配置文件，自动合并默认值"""
        for name, filename in CONFIG_FILES.items():
            # 获取用户配置（可能为空）
            user_config = self._load_single_config(filename)
            
            # 获取对应的默认配置
            default_config = DEFAULT_CONFIG.get(name, {})
            
            # 深度合并配置（用户配置优先）
            merged_config = self._deep_merge(default_config, user_config)
            self._configs[name] = merged_config
            
            safe_log.info("加载配置块: %s | 文件: %s", name, filename)
    
    def _load_single_config(self, filename: str) -> Dict[str, Any]:
        """安全加载单个配置文件"""
        filepath = os.path.join(self._config_dir, filename)
        
        if not os.path.exists(filepath):
            safe_log.debug("配置文件不存在: %s", filepath)
            return {}
        
        user_config = read_toml_file(filepath)
        if not user_config:
            safe_log.warning("配置文件为空或格式错误: %s", filepath)
            return {}
        
        return user_config
    
    def _deep_merge(self, default: Dict, user: Dict) -> Dict:
        """
        深度合并两个字典，用户配置优先
        
        参数:
            default: 默认配置
            user: 用户配置
            
        返回:
            Dict: 合并后的配置
        """
        result = default.copy()
        
        for key, user_value in user.items():
            if key not in result:
                # 新键：直接添加
                result[key] = user_value
            elif isinstance(result[key], dict) and isinstance(user_value, dict):
                # 字典：递归合并
                result[key] = self._deep_merge(result[key], user_value)
            else:
                # 其他类型：用户配置优先
                result[key] = user_value
        
        return result
    
    # --------------------------
    # 配置文件路径获取方法
    # --------------------------
    def get_config_file_path(self, config_name: str) -> str:
        """
        获取配置文件完整路径
        
        参数:
            config_name: 配置名称（如 "proofread_symbols", "gui_config" 等）
            
        返回:
            str: 配置文件完整路径
            
        异常:
            ValueError: 未知的配置名称
        """
        if config_name not in CONFIG_FILES:
            raise ValueError(f"未知的配置名称: {config_name}")
        filename = CONFIG_FILES[config_name]
        return os.path.join(self._config_dir, filename)
    
    # --------------------------
    # 第一层：配置块级别访问方法
    # --------------------------
    def get_logger_config_block(self) -> Dict[str, Any]:
        """获取整个日志配置块"""
        return self._configs.get("logger_config", {})
    
    # 校对配置块
    def get_proofread_config_block(self) -> Dict[str, Any]:
        """获取校对主配置块"""
        return self._configs.get("proofread_config", {})
    
    def get_proofread_symbols_block(self) -> Dict[str, Any]:
        """获取符号校对配置块"""
        return self._configs.get("proofread_symbols", {})

    def get_proofread_typos_block(self) -> Dict[str, Any]:
        """获取错别字配置块"""
        return self._configs.get("proofread_typos", {})

    def get_proofread_sensitive_block(self) -> Dict[str, Any]:
        """获取敏感词配置块"""
        return self._configs.get("proofread_sensitive", {})
    
    def get_gui_config_block(self) -> Dict[str, Any]:
        """获取整个GUI配置块"""
        return self._configs.get("gui_config", {})

    def get_link_config_block(self) -> Dict[str, Any]:
        """获取整个链接配置块"""
        return self._configs.get("link_config", {})
    
    def get_output_config_block(self) -> Dict[str, Any]:
        """获取整个输出配置块"""
        return self._configs.get("output_config", {})
    
    def get_software_priority_block(self) -> Dict[str, Any]:
        """获取整个软件优先级配置块"""
        return self._configs.get("software_priority", {})
    
    def get_conversion_defaults_block(self) -> Dict[str, Any]:
        """获取整个转换默认值配置块（控制 GUI 界面的默认值设置）"""
        return self._configs.get("conversion_defaults", {})
    
    def get_heading_numbering_config_block(self) -> Dict[str, Any]:
        """获取整个标题序号配置块"""
        return self._configs.get("heading_numbering_add", {})
    
    def get_numbering_patterns_block(self) -> Dict[str, Any]:
        """获取整个序号清理规则配置块"""
        # 优先从配置文件加载，否则使用默认配置
        config = self._configs.get("heading_numbering_clean", {})
        if not config:
            config = DEFAULT_NUMBERING_PATTERNS_CONFIG.get("heading_numbering_clean", {})
        return config
    
    def get_style_code_block(self) -> Dict[str, Any]:
        """获取代码样式配置块"""
        return self._configs.get("style_code", {})
    
    def get_style_quote_block(self) -> Dict[str, Any]:
        """获取引用样式配置块"""
        return self._configs.get("style_quote", {})
    
    def get_style_formula_block(self) -> Dict[str, Any]:
        """获取公式样式配置块"""
        return self._configs.get("style_formula", {})
    
    def get_style_table_block(self) -> Dict[str, Any]:
        """获取表格样式配置块"""
        return self._configs.get("style_table", {})
    
    # --------------------------
    # 序号清理规则相关方法
    # --------------------------
    
    def get_cleaning_rules(self) -> List[Dict[str, Any]]:
        """
        获取启用的清理规则列表（按优先级排序，已替换占位符）
        
        返回:
            List[Dict]: 规则列表，每个规则包含:
                - rule_id: str - 规则ID
                - name: str - 规则名称
                - description: str - 规则说明
                - enabled: bool - 是否启用
                - is_system: bool - 是否系统规则
                - regex: str - 原始正则表达式
                - compiled_regex: str - 替换占位符后的正则表达式
        """
        # 从 heading_utils.py 导入占位符定义（硬编码，使用 raw string）
        from gongwen_converter.utils.heading_utils import NUMBERING_PLACEHOLDERS
        
        config = self.get_numbering_patterns_block()
        placeholders = NUMBERING_PLACEHOLDERS  # 使用代码中的占位符
        settings = config.get("settings", {})
        rules_dict = config.get("rules", {})
        order = settings.get("order", list(rules_dict.keys()))
        
        result = []
        for rule_id in order:
            if rule_id not in rules_dict:
                continue
            
            rule = rules_dict[rule_id]
            if not rule.get("enabled", True):
                continue
            
            # 替换占位符
            regex = rule.get("regex", "")
            compiled_regex = self._replace_placeholders(regex, placeholders)
            
            result.append({
                "rule_id": rule_id,
                "name": rule.get("name", rule_id),
                "description": rule.get("description", ""),
                "enabled": rule.get("enabled", True),
                "is_system": rule.get("is_system", False),
                "regex": regex,
                "compiled_regex": compiled_regex
            })
        
        safe_log.debug("获取清理规则列表，共 %d 条启用的规则", len(result))
        return result
    
    def _replace_placeholders(self, regex: str, placeholders: Dict[str, str]) -> str:
        """
        替换正则表达式中的占位符
        
        参数:
            regex: 原始正则表达式（含 {placeholder} 格式的占位符）
            placeholders: 占位符字典
            
        返回:
            str: 替换后的正则表达式
        """
        result = regex
        for name, value in placeholders.items():
            result = result.replace(f"{{{name}}}", value)
        return result
    
    # --------------------------
    # 第二层：子表级别访问方法
    # --------------------------
    
    # 日志配置子表
    def get_logging_config(self) -> Dict[str, Any]:
        """获取日志配置子表"""
        return self.get_logger_config_block().get("logging", {})
    
    # 校对配置子表
    def get_proofread_engine_config(self) -> Dict[str, Any]:
        """获取校对引擎配置（开关和跳过规则）"""
        return self.get_proofread_config_block().get("engine", {})

    def get_symbol_pairing_config(self) -> Dict[str, Any]:
        """获取符号配对子表"""
        return self.get_proofread_symbols_block().get("symbol_pairing", {})
    
    def get_symbol_map_config(self) -> Dict[str, Any]:
        """获取符号映射子表"""
        return self.get_proofread_symbols_block().get("symbol_map", {})
    
    def get_typos_config(self) -> Dict[str, Any]:
        """获取错别字子表"""
        return self.get_proofread_typos_block().get("typos", {})

    def get_sensitive_words_config(self) -> Dict[str, Any]:
        """获取敏感词子表"""
        return self.get_proofread_sensitive_block().get("sensitive_words", {})
    
    # 校对跳过规则方法
    def is_skip_code_blocks_enabled(self) -> bool:
        """
        是否跳过代码块段落
        
        返回:
            bool: 是否跳过
        """
        config = self.get_proofread_engine_config()
        enabled = config.get("skip_code_blocks", True)
        safe_log.debug("校对跳过代码块: %s", enabled)
        return enabled
    
    def is_skip_quote_blocks_enabled(self) -> bool:
        """
        是否跳过引用段落
        
        返回:
            bool: 是否跳过
        """
        config = self.get_proofread_engine_config()
        enabled = config.get("skip_quote_blocks", False)
        safe_log.debug("校对跳过引用块: %s", enabled)
        return enabled
    
    # GUI配置子表
    def get_window_config(self) -> Dict[str, Any]:
        """获取窗口设置子表"""
        return self.get_gui_config_block().get("window", {})

    def get_component_config(self) -> Dict[str, Any]:
        """获取组件设置子表"""
        return self.get_gui_config_block().get("component", {})
    
    def get_theme_config(self) -> Dict[str, Any]:
        """获取主题设置子表"""
        return self.get_gui_config_block().get("theme", {})
    
    def get_transparency_config(self) -> Dict[str, Any]:
        """获取透明度设置子表"""
        return self.get_gui_config_block().get("transparency", {})

    def get_template_config(self) -> Dict[str, Any]:
        """获取模板设置子表（从GUI配置）"""
        return self.get_gui_config_block().get("template", {})
    
    def get_language_config(self) -> Dict[str, Any]:
        """获取语言设置子表（从GUI配置）"""
        return self.get_gui_config_block().get("language", {})
    
    def get_link_format_config(self) -> Dict[str, Any]:
        """获取链接格式配置子表"""
        return self.get_link_config_block().get("format", {})
    
    def get_non_embed_links_config(self) -> Dict[str, Any]:
        """获取非嵌入链接配置子表"""
        return self.get_link_config_block().get("non_embed_links", {})
    
    def get_embed_links_config(self) -> Dict[str, Any]:
        """获取嵌入链接配置子表"""
        return self.get_link_config_block().get("embed_links", {})
    
    def get_embedding_config(self) -> Dict[str, Any]:
        """获取嵌入详细配置子表"""
        return self.get_link_config_block().get("embedding", {})
    
    def get_path_resolution_config(self) -> Dict[str, Any]:
        """获取路径解析配置子表"""
        return self.get_link_config_block().get("path_resolution", {})
    
    def get_link_error_handling_config(self) -> Dict[str, Any]:
        """获取链接错误处理配置子表"""
        return self.get_link_config_block().get("error_handling", {})
    
    # Conversion Defaults 配置子表
    def get_document_defaults(self) -> Dict[str, Any]:
        """获取文档文件默认设置子表"""
        return self.get_conversion_defaults_block().get("document", {})
    
    def get_spreadsheet_defaults(self) -> Dict[str, Any]:
        """获取表格文件默认设置子表"""
        return self.get_conversion_defaults_block().get("spreadsheet", {})
    
    def get_image_defaults(self) -> Dict[str, Any]:
        """获取图片文件默认设置子表"""
        return self.get_conversion_defaults_block().get("image", {})
    
    def get_layout_defaults(self) -> Dict[str, Any]:
        """获取版式文件默认设置子表"""
        return self.get_conversion_defaults_block().get("layout", {})
    
    def get_text_defaults(self) -> Dict[str, Any]:
        """获取文本文件默认设置子表"""
        return self.get_conversion_defaults_block().get("text", {})

    # Output配置子表
    
    def get_output_directory_settings(self) -> Dict[str, Any]:
        """获取输出目录设置子表"""
        return self.get_output_config_block().get("directory", {})
    
    def get_output_behavior_settings(self) -> Dict[str, Any]:
        """获取输出行为设置子表"""
        return self.get_output_config_block().get("behavior", {})
    
    def get_output_intermediate_files_config(self) -> Dict[str, Any]:
        """获取输出中间文件配置子表（从输出配置）"""
        return self.get_output_config_block().get("intermediate_files", {})
    
    # 软件优先级配置子表
    def get_default_priority_config(self) -> Dict[str, Any]:
        """获取默认优先级配置子表"""
        return self.get_software_priority_block().get("default_priority", {})

    def get_special_conversions_config(self) -> Dict[str, Any]:
        """获取特殊转换配置子表"""
        return self.get_software_priority_block().get("special_conversions", {})
    
    # --------------------------
    # 第三层：具体配置获取方法
    # --------------------------
    
    def get_typos(self) -> Dict[str, list]:
        """获取错别字映射表（带默认值）"""
        typos = self.get_typos_config()
        return typos if isinstance(typos, dict) else {}
    
    def get_sensitive_words(self) -> Dict[str, list]:
        """获取敏感词映射表（带默认值）"""
        words = self.get_sensitive_words_config()
        return words if isinstance(words, dict) else {}
    
    def get_symbol_map(self) -> Dict[str, list]:
        """获取标点符号映射（带默认值，键统一为字符串）"""
        symbol_map = self.get_symbol_map_config()
        if isinstance(symbol_map, dict):
            # 统一将所有键转为字符串，兼容整数键和字符串键
            # 这确保 GUI 编辑后的字符串键与配置文件中的数字键行为一致
            return {str(k): v for k, v in symbol_map.items()}
        return {}
    
    def get_symbol_pairs(self) -> List[Tuple[str, str]]:
        """获取符号对列表（带默认值）"""
        pairing_config = self.get_symbol_pairing_config()
        pairs = pairing_config.get("pairs", [])
        
        # 转换为元组列表
        if isinstance(pairs, list):
            return [tuple(pair) for pair in pairs if isinstance(pair, list) and len(pair) == 2]
        return []
    
    def get_center_panel_width(self) -> int:
        """
        获取中栏宽度
        
        返回:
            int: 中栏宽度
        """
        window_config = self.get_window_config()
        width = window_config.get("center_panel_width", 400)
        safe_log.debug("获取中栏宽度: %d", width)
        return width
    
    def get_batch_panel_width(self) -> int:
        """
        获取批量面板宽度

        返回:
            int: 批量面板宽度
        """
        window_config = self.get_window_config()
        width = window_config.get("left_panel_width", 400)
        safe_log.debug("获取批量面板宽度: %d", width)
        return width

    def get_template_panel_width(self) -> int:
        """
        获取模板面板宽度

        返回:
            int: 模板面板宽度
        """
        window_config = self.get_window_config()
        width = window_config.get("right_panel_width", 300)
        safe_log.debug("获取模板面板宽度: %d", width)
        return width
    
    def get_center_panel_screen_x(self) -> int:
        """
        获取中栏在屏幕上的X坐标
        
        返回:
            int: 中栏屏幕X坐标
        """
        window_config = self.get_window_config()
        x = window_config.get("center_panel_screen_x", 0)
        safe_log.debug("获取中栏屏幕X坐标: %d", x)
        return x
    
    def get_default_mode(self) -> str:
        """
        获取默认启动模式
        
        返回:
            str: 默认模式，"single" 或 "batch"
        """
        window_config = self.get_window_config()
        mode = window_config.get("default_mode", "single")
        if mode not in ["single", "batch"]:
            mode = "single"
        safe_log.debug("获取默认启动模式: %s", mode)
        return mode
    
    def get_window_height(self) -> int:
        """
        获取窗口默认高度
        
        返回:
            int: 窗口高度
        """
        window_config = self.get_window_config()
        height = window_config.get("default_height", 740)
        safe_log.debug("获取窗口高度: %d", height)
        return height

    def get_min_height(self) -> int:
        """
        获取窗口最小高度
        
        返回:
            int: 最小高度
        """
        window_config = self.get_window_config()
        min_height = window_config.get("min_height", 720)
        safe_log.debug("获取窗口最小高度: %d", min_height)
        return min_height
    
    def get_file_drop_height(self) -> int:
        """
        获取文件拖拽区域高度
        
        返回:
            int: 文件拖拽区域高度
        """
        component_config = self.get_component_config()
        height = component_config.get("file_drop_height")
        safe_log.debug("获取文件拖拽区域高度: %d", height)
        return height

    def get_window_y(self) -> int:
        """
        获取窗口Y坐标
        
        返回:
            int: 窗口Y坐标
        """
        window_config = self.get_window_config()
        y = window_config.get("window_y", 0)
        safe_log.debug("获取窗口Y坐标: %d", y)
        return y

    def get_window_position(self) -> Tuple[int, int]:
        """
        获取窗口默认位置（基于中栏屏幕坐标计算）
        
        返回:
            Tuple[int, int]: (x坐标, y坐标)
        """
        window_config = self.get_window_config()
        
        # 获取中栏屏幕坐标和窗口Y坐标
        center_x = window_config.get("center_panel_screen_x", 0)
        y = window_config.get("window_y", 0)
        
        # 单文件模式：窗口X = 中栏X（中栏居中显示）
        window_x = center_x
        
        safe_log.debug("获取窗口位置: (%d, %d) [从中栏坐标%d计算]", window_x, y, center_x)
        return window_x, y

    def get_default_theme(self) -> str:
        """
        获取默认主题名称
        
        返回:
            str: 主题名称
        """
        theme_config = self.get_theme_config()
        theme = theme_config.get("default_theme")
        safe_log.debug("获取默认主题: %s", theme)
        return theme

    def is_transparency_enabled(self) -> bool:
        """
        检查是否启用透明度效果
        
        返回:
            bool: 是否启用透明度
        """
        transparency_config = self.get_transparency_config()
        enabled = transparency_config.get("enabled")
        safe_log.debug("透明度启用状态: %s", enabled)
        return enabled

    def get_transparency_value(self) -> float:
        """
        获取透明度值
        
        返回:
            float: 透明度值 (0.0-1.0)
        """
        transparency_config = self.get_transparency_config()
        transparency = transparency_config.get("default_value")
        # 确保在有效范围内
        transparency = max(0.1, min(1.0, transparency))
        safe_log.debug("获取透明度值: %.2f", transparency)
        return transparency

    def should_remember_gui_state(self) -> bool:
        """
        检查是否应记住窗口位置
        
        返回:
            bool: 是否记住窗口位置
        """
        window_config = self.get_window_config()
        remember = window_config.get("remember_gui_state")
        safe_log.debug("记住窗口位置: %s", remember)
        return remember

    def should_auto_center(self) -> bool:
        """
        检查是否应自动居中窗口
        
        返回:
            bool: 是否自动居中窗口
        """
        window_config = self.get_window_config()
        auto_center = window_config.get("auto_center")
        safe_log.debug("自动居中窗口: %s", auto_center)
        return auto_center

    def get_default_md_template_type(self) -> str:
        """
        获取默认MD文件模板类型（从GUI配置）
        
        返回:
            str: 默认模板类型 ("docx" 或 "xlsx")
        """
        template_config = self.get_template_config()
        template_type = template_config.get("md_default_template", "docx")
        # 确保返回有效值
        if template_type not in ["docx", "xlsx"]:
            template_type = "docx"
        safe_log.debug("获取默认MD模板类型: %s", template_type)
        return template_type
    
    def get_locale(self) -> str:
        """
        获取当前语言设置（从GUI配置）
        
        返回:
            str: 语言代码，如 "zh_CN" 或 "en_US"
        """
        language_config = self.get_language_config()
        locale = language_config.get("locale", "zh_CN")
        safe_log.debug("获取语言设置: %s", locale)
        return locale
    
    def get_docx_to_md_keep_images(self) -> bool:
        """
        获取文档转MD时是否保留图片
        
        返回:
            bool: 是否保留图片
        """
        document_config = self.get_document_defaults()
        keep_images = document_config.get("to_md_keep_images", True)
        safe_log.debug("文档转MD保留图片: %s", keep_images)
        return keep_images
    
    def get_docx_to_md_enable_ocr(self) -> bool:
        """
        获取文档转MD时是否启用OCR
        
        返回:
            bool: 是否启用OCR
        """
        document_config = self.get_document_defaults()
        enable_ocr = document_config.get("to_md_enable_ocr", False)
        safe_log.debug("文档转MD启用OCR: %s", enable_ocr)
        return enable_ocr
    
    def get_xlsx_to_md_keep_images(self) -> bool:
        """
        获取表格转MD时是否保留图片
        
        返回:
            bool: 是否保留图片
        """
        spreadsheet_config = self.get_spreadsheet_defaults()
        keep_images = spreadsheet_config.get("to_md_keep_images", True)
        safe_log.debug("表格转MD保留图片: %s", keep_images)
        return keep_images
    
    def get_xlsx_to_md_enable_ocr(self) -> bool:
        """
        获取表格转MD时是否启用OCR
        
        返回:
            bool: 是否启用OCR
        """
        spreadsheet_config = self.get_spreadsheet_defaults()
        enable_ocr = spreadsheet_config.get("to_md_enable_ocr", False)
        safe_log.debug("表格转MD启用OCR: %s", enable_ocr)
        return enable_ocr
    
    def get_layout_to_md_keep_images(self) -> bool:
        """
        获取版式文件转MD时是否保留图片
        
        返回:
            bool: 是否保留图片
        """
        layout_config = self.get_layout_defaults()
        keep_images = layout_config.get("to_md_keep_images", True)
        safe_log.debug("版式转MD保留图片: %s", keep_images)
        return keep_images
    
    def get_layout_to_md_enable_ocr(self) -> bool:
        """
        获取版式文件转MD时是否启用OCR
        
        返回:
            bool: 是否启用OCR
        """
        layout_config = self.get_layout_defaults()
        enable_ocr = layout_config.get("to_md_enable_ocr", False)
        safe_log.debug("版式转MD启用OCR: %s", enable_ocr)
        return enable_ocr
    
    def get_image_to_md_keep_images(self) -> bool:
        """
        获取图片文件转MD时是否保留图片
        
        返回:
            bool: 是否保留图片
        """
        image_config = self.get_image_defaults()
        keep_images = image_config.get("to_md_keep_images", True)
        safe_log.debug("图片转MD保留图片: %s", keep_images)
        return keep_images
    
    def get_image_to_md_enable_ocr(self) -> bool:
        """
        获取图片文件转MD时是否启用OCR
        
        返回:
            bool: 是否启用OCR
        """
        image_config = self.get_image_defaults()
        enable_ocr = image_config.get("to_md_enable_ocr", True)
        safe_log.debug("图片转MD启用OCR: %s", enable_ocr)
        return enable_ocr
    
    # File Defaults 新增的具体配置获取方法
    def get_document_compress_mode(self) -> str:
        """获取文档压缩模式（暂未使用）"""
        document_config = self.get_document_defaults()
        return document_config.get("compress_mode", "lossless")
    
    def get_spreadsheet_merge_mode(self) -> int:
        """获取表格汇总模式默认值"""
        spreadsheet_config = self.get_spreadsheet_defaults()
        mode = spreadsheet_config.get("merge_mode", 3)
        safe_log.debug("表格汇总模式: %d", mode)
        return mode
    
    def get_image_compress_mode(self) -> str:
        """获取图片压缩模式默认值"""
        image_config = self.get_image_defaults()
        mode = image_config.get("compress_mode", "lossless")
        safe_log.debug("图片压缩模式: %s", mode)
        return mode
    
    def get_image_size_limit(self) -> int:
        """获取图片大小限制默认值"""
        image_config = self.get_image_defaults()
        limit = image_config.get("size_limit", 200)
        safe_log.debug("图片大小限制: %d", limit)
        return limit
    
    def get_image_size_unit(self) -> str:
        """获取图片大小单位默认值"""
        image_config = self.get_image_defaults()
        unit = image_config.get("size_unit", "KB")
        safe_log.debug("图片大小单位: %s", unit)
        return unit
    
    def get_image_pdf_quality(self) -> str:
        """获取图片转PDF质量默认值"""
        image_config = self.get_image_defaults()
        quality = image_config.get("pdf_quality", "original")
        safe_log.debug("图片转PDF质量: %s", quality)
        return quality
    
    def get_image_tiff_mode(self) -> str:
        """获取图片转TIFF模式默认值"""
        image_config = self.get_image_defaults()
        mode = image_config.get("tiff_mode", "smart")
        safe_log.debug("图片转TIFF模式: %s", mode)
        return mode
    
    def get_layout_render_dpi(self) -> int:
        """获取版式文件渲染DPI默认值"""
        layout_config = self.get_layout_defaults()
        dpi = layout_config.get("render_dpi", 300)
        safe_log.debug("版式渲染DPI: %d", dpi)
        return dpi
    
    def get_document_enable_symbol_pairing(self) -> bool:
        """获取文档校对-标点配对默认值"""
        document_config = self.get_document_defaults()
        enabled = document_config.get("enable_symbol_pairing", True)
        safe_log.debug("文档标点配对: %s", enabled)
        return enabled
    
    def get_document_enable_symbol_correction(self) -> bool:
        """获取文档校对-符号校对默认值"""
        document_config = self.get_document_defaults()
        enabled = document_config.get("enable_symbol_correction", True)
        safe_log.debug("文档符号校对: %s", enabled)
        return enabled
    
    def get_document_enable_typos_rule(self) -> bool:
        """获取文档校对-错别字默认值"""
        document_config = self.get_document_defaults()
        enabled = document_config.get("enable_typos_rule", True)
        safe_log.debug("文档错别字校对: %s", enabled)
        return enabled
    
    def get_document_enable_sensitive_word(self) -> bool:
        """获取文档校对-敏感词默认值"""
        document_config = self.get_document_defaults()
        enabled = document_config.get("enable_sensitive_word", True)
        safe_log.debug("文档敏感词匹配: %s", enabled)
        return enabled
    
    # 标题序号配置获取方法（文档转MD）
    def get_docx_to_md_remove_numbering(self) -> bool:
        """
        获取文档转MD时是否默认清除原序号
        
        返回:
            bool: 是否清除原序号
        """
        document_config = self.get_document_defaults()
        remove = document_config.get("to_md_remove_numbering", True)
        safe_log.debug("文档转MD清除原序号: %s", remove)
        return remove
    
    def get_docx_to_md_add_numbering(self) -> bool:
        """
        获取文档转MD时是否默认添加新序号
        
        返回:
            bool: 是否添加新序号
        """
        document_config = self.get_document_defaults()
        add = document_config.get("to_md_add_numbering", False)
        safe_log.debug("文档转MD添加新序号: %s", add)
        return add
    
    def get_docx_to_md_default_scheme(self) -> str:
        """
        获取文档转MD的默认序号方案
        
        返回:
            str: 序号方案名称，如果配置的方案不存在则返回全局默认方案
        """
        try:
            document_config = self.get_document_defaults()
            scheme_name = document_config.get("to_md_default_scheme", "gongwen_standard")
            
            # 验证方案是否存在
            heading_config = self.get_heading_numbering_config_block()
            schemes = heading_config.get("schemes", {})
            
            if scheme_name not in schemes:
                safe_log.warning(
                    "配置的序号方案 '%s' 不存在，使用全局默认方案 'gongwen_standard'",
                    scheme_name
                )
                return "gongwen_standard"
            
            safe_log.debug("文档转MD默认序号方案: %s", scheme_name)
            return scheme_name
        except Exception as e:
            safe_log.error("获取文档转MD默认序号方案失败: %s", str(e))
            return "gongwen_standard"
    
    def get_docx_to_md_enable_optimization(self) -> bool:
        """
        获取文档转MD时是否默认启用针对性优化
        
        返回:
            bool: 是否启用针对文档类型优化
        """
        document_config = self.get_document_defaults()
        enable = document_config.get("to_md_enable_optimization", True)
        safe_log.debug("文档转MD启用优化: %s", enable)
        return enable
    
    def get_docx_to_md_optimization_type(self) -> str:
        """
        获取文档转MD的默认优化类型
        
        返回:
            str: 优化类型（"公文"/"合同"/"论文"）
        """
        document_config = self.get_document_defaults()
        opt_type = document_config.get("to_md_optimization_type", "公文")
        safe_log.debug("文档转MD优化类型: %s", opt_type)
        return opt_type
    
    # 标题序号配置获取方法（MD转文档）
    def get_md_to_docx_remove_numbering(self) -> bool:
        """
        获取MD转DOCX时是否默认清除原序号
        
        返回:
            bool: 是否清除原序号
        """
        text_config = self.get_text_defaults()
        remove = text_config.get("to_docx_remove_numbering", True)
        safe_log.debug("MD转DOCX清除原序号: %s", remove)
        return remove
    
    def get_md_to_docx_add_numbering(self) -> bool:
        """
        获取MD转DOCX时是否默认添加新序号
        
        返回:
            bool: 是否添加新序号
        """
        text_config = self.get_text_defaults()
        add = text_config.get("to_docx_add_numbering", True)
        safe_log.debug("MD转DOCX添加新序号: %s", add)
        return add
    
    def get_md_to_docx_default_scheme(self) -> str:
        """
        获取MD转DOCX的默认序号方案
        
        返回:
            str: 序号方案名称，如果配置的方案不存在则返回全局默认方案
        """
        try:
            text_config = self.get_text_defaults()
            scheme_name = text_config.get("to_docx_default_scheme", "gongwen_standard")
            
            # 验证方案是否存在
            heading_config = self.get_heading_numbering_config_block()
            schemes = heading_config.get("schemes", {})
            
            if scheme_name not in schemes:
                safe_log.warning(
                    "配置的序号方案 '%s' 不存在，使用全局默认方案 'gongwen_standard'",
                    scheme_name
                )
                return "gongwen_standard"
            
            safe_log.debug("MD转DOCX默认序号方案: %s", scheme_name)
            return scheme_name
        except Exception as e:
            safe_log.error("获取MD转DOCX默认序号方案失败: %s", str(e))
            return "gongwen_standard"
    
    # 序号方案相关方法
    def get_heading_schemes(self) -> Dict[str, Dict]:
        """
        获取所有序号方案
        
        返回:
            Dict[str, Dict]: 方案字典，键为方案ID，值为方案配置
        """
        heading_config = self.get_heading_numbering_config_block()
        schemes = heading_config.get("schemes", {})
        safe_log.debug("获取序号方案列表，共 %d 个方案", len(schemes))
        return schemes
    
    def get_scheme_names(self) -> List[str]:
        """
        获取所有序号方案的名称列表（供GUI下拉框使用）
        
        返回:
            List[str]: 方案显示名称列表，如 ["公文标准", "层级数字标准", "法律条文标准"]
        """
        schemes = self.get_heading_schemes()
        scheme_names = []
        
        # 提取每个方案的name字段
        for scheme_id, scheme_config in schemes.items():
            if isinstance(scheme_config, dict) and "name" in scheme_config:
                scheme_names.append(scheme_config["name"])
            else:
                # 如果没有name字段，使用ID
                scheme_names.append(scheme_id)
        
        safe_log.debug("获取序号方案名称列表: %s", scheme_names)
        return scheme_names
    
    def get_save_intermediate_files(self) -> bool:
        """
        获取是否保存中间文件到输出目录（从输出配置）
        
        返回:
            bool: 是否保存中间文件
        """
        intermediate_config = self.get_output_intermediate_files_config()
        save_to_output = intermediate_config.get("save_to_output", True)
        safe_log.debug("保存中间文件设置: %s", save_to_output)
        return save_to_output
    
    def get_markdown_link_style_settings(self) -> Dict[str, Any]:
        """
        获取Markdown链接格式设置（从链接配置）
        
        返回:
            Dict[str, Any]: 包含以下键的字典：
                - image_link_style: str - 图片链接样式 
                    可选值: "markdown_embed", "markdown_link", "wiki_embed", "wiki_link"
                - md_file_link_style: str - MD文件链接样式
                    可选值: "markdown_link", "wiki_embed", "wiki_link"
        """
        link_format_config = self.get_link_format_config()
        settings = {
            "image_link_style": link_format_config.get("image_link_style", "wiki_embed"),
            "md_file_link_style": link_format_config.get("md_file_link_style", "wiki_embed")
        }
        safe_log.debug("获取Markdown链接格式设置: %s", settings)
        return settings
    
    def get_image_link_style(self) -> str:
        """
        获取图片链接样式
        
        返回:
            str: 样式值 ("markdown_embed", "markdown_link", "wiki_embed", "wiki_link")
        """
        config = self.get_link_format_config()
        style = config.get("image_link_style", "wiki_embed")
        safe_log.debug("获取图片链接样式: %s", style)
        return style
    
    def get_md_file_link_style(self) -> str:
        """
        获取MD文件链接样式
        
        返回:
            str: 样式值 ("markdown_link", "wiki_embed", "wiki_link")
        """
        config = self.get_link_format_config()
        style = config.get("md_file_link_style", "wiki_embed")
        safe_log.debug("获取MD文件链接样式: %s", style)
        return style
    
    # 非嵌入链接配置获取方法
    def get_wiki_link_mode(self) -> str:
        """
        获取Wiki链接处理模式
        
        返回:
            str: 处理模式 ("keep", "extract_text", "remove")
        """
        config = self.get_non_embed_links_config()
        mode = config.get("wiki_mode", "extract_text")
        safe_log.debug("获取Wiki链接处理模式: %s", mode)
        return mode
    
    def get_markdown_link_mode(self) -> str:
        """
        获取Markdown链接处理模式
        
        返回:
            str: 处理模式 ("keep", "extract_text", "remove")
        """
        config = self.get_non_embed_links_config()
        mode = config.get("markdown_mode", "extract_text")
        safe_log.debug("获取Markdown链接处理模式: %s", mode)
        return mode
    
    # 嵌入链接配置获取方法
    def get_wiki_embed_image_mode(self) -> str:
        """
        获取Wiki嵌入图片处理模式（![[image.png]]）
        
        返回:
            str: 处理模式 ("keep", "extract_text", "remove", "embed")
        """
        config = self.get_embed_links_config()
        mode = config.get("wiki_image_mode", "embed")
        safe_log.debug("获取Wiki嵌入图片处理模式: %s", mode)
        return mode
    
    def get_markdown_embed_image_mode(self) -> str:
        """
        获取Markdown嵌入图片处理模式（![alt](image.png)）
        
        返回:
            str: 处理模式 ("keep", "extract_text", "remove", "embed")
        """
        config = self.get_embed_links_config()
        mode = config.get("markdown_image_mode", "embed")
        safe_log.debug("获取Markdown嵌入图片处理模式: %s", mode)
        return mode
    
    def get_embed_md_file_mode(self) -> str:
        """
        获取嵌入MD文件处理模式
        
        返回:
            str: 处理模式 ("keep", "extract_text", "remove", "embed")
        """
        config = self.get_embed_links_config()
        mode = config.get("md_file_mode", "embed")
        safe_log.debug("获取嵌入MD文件处理模式: %s", mode)
        return mode
    
    # 嵌入详细配置获取方法
    def get_max_embed_depth(self) -> int:
        """
        获取最大嵌入深度
        
        返回:
            int: 最大递归嵌入深度
        """
        config = self.get_embedding_config()
        depth = config.get("max_depth", 3)
        safe_log.debug("获取最大嵌入深度: %d", depth)
        return depth
    
    # 路径解析配置获取方法
    def get_search_dirs(self) -> List[str]:
        """
        获取搜索目录列表
        
        返回:
            List[str]: 文件搜索子目录列表
        """
        config = self.get_path_resolution_config()
        dirs = config.get("search_dirs", [".", "assets", "images", "attachments"])
        safe_log.debug("获取搜索目录列表: %s", dirs)
        return dirs
    
    # 错误处理配置获取方法
    def get_file_not_found_mode(self) -> str:
        """
        获取文件未找到处理模式
        
        返回:
            str: 处理模式 ("ignore", "keep", "placeholder")
        """
        config = self.get_link_error_handling_config()
        mode = config.get("file_not_found", "placeholder")
        safe_log.debug("获取文件未找到处理模式: %s", mode)
        return mode
    
    def get_file_not_found_text(self) -> str:
        """
        获取文件未找到占位文本
        
        返回:
            str: 占位文本模板，支持 {filename} 占位符
        """
        config = self.get_link_error_handling_config()
        text = config.get("file_not_found_text", "⚠️ 文件未找到: {filename}")
        safe_log.debug("获取文件未找到占位文本: %s", text)
        return text
    
    def is_circular_detection_enabled(self) -> bool:
        """
        是否启用循环引用检测
        
        返回:
            bool: 是否检测循环引用
        """
        config = self.get_link_error_handling_config()
        enabled = config.get("detect_circular", True)
        safe_log.debug("循环引用检测启用状态: %s", enabled)
        return enabled
    
    def get_circular_reference_mode(self) -> str:
        """
        获取循环引用处理模式
        
        返回:
            str: 处理模式 ("ignore", "keep", "placeholder")
        """
        config = self.get_link_error_handling_config()
        mode = config.get("circular_reference", "placeholder")
        safe_log.debug("获取循环引用处理模式: %s", mode)
        return mode
    
    def get_circular_text(self) -> str:
        """
        获取循环引用占位文本
        
        返回:
            str: 占位文本模板，支持 {filename} 占位符
        """
        config = self.get_link_error_handling_config()
        text = config.get("circular_text", "⚠️ 检测到循环引用: {filename}")
        safe_log.debug("获取循环引用占位文本: %s", text)
        return text
    
    def get_max_depth_reached_mode(self) -> str:
        """
        获取达到最大深度处理模式
        
        返回:
            str: 处理模式 ("ignore", "keep", "placeholder")
        """
        config = self.get_link_error_handling_config()
        mode = config.get("max_depth_reached", "placeholder")
        safe_log.debug("获取最大深度处理模式: %s", mode)
        return mode
    
    def get_max_depth_text(self) -> str:
        """
        获取最大深度警告文本
        
        返回:
            str: 警告文本
        """
        config = self.get_link_error_handling_config()
        text = config.get("max_depth_text", "⚠️ 达到最大嵌入深度")
        safe_log.debug("获取最大深度警告文本: %s", text)
        return text
    
    def get_auto_open_folder(self) -> bool:
        """
        获取是否自动打开输出文件夹
        
        返回:
            bool: 是否自动打开
        """
        behavior_settings = self.get_output_behavior_settings()
        auto_open = behavior_settings.get("auto_open_folder", False)
        safe_log.debug("自动打开输出文件夹: %s", auto_open)
        return auto_open

    # 软件优先级具体配置获取方法
    def get_word_processors_priority(self) -> List[str]:
        """获取文档处理软件优先级列表"""
        default_priority = self.get_default_priority_config()
        return default_priority.get("word_processors", ["wps_writer", "msoffice_word", "libreoffice"])

    def get_spreadsheet_processors_priority(self) -> List[str]:
        """获取表格处理软件优先级列表"""
        default_priority = self.get_default_priority_config()
        return default_priority.get("spreadsheet_processors", ["wps_spreadsheets", "msoffice_excel", "libreoffice"])

    def get_special_conversion_priority(self, conversion_type: str) -> List[str]:
        """获取特定特殊转换的软件优先级列表"""
        special_conversions = self.get_special_conversions_config()
        return special_conversions.get(conversion_type, [])
    
    # --------------------------
    # 转换行为配置方法
    # --------------------------
    
    def get_conversion_config_block(self) -> Dict[str, Any]:
        """获取整个转换行为配置块（直接控制转换引擎的行为规则）"""
        return self._configs.get("conversion_config", {})
    
    def get_docx_to_md_config(self) -> Dict[str, Any]:
        """获取DOCX转MD配置子表"""
        return self.get_conversion_config_block().get("docx_to_md", {})
    
    def get_md_to_docx_config(self) -> Dict[str, Any]:
        """获取MD转DOCX配置子表"""
        return self.get_conversion_config_block().get("md_to_docx", {})
    
    def get_formatting_syntax_config(self) -> Dict[str, Any]:
        """获取Markdown语法配置子表"""
        return self.get_conversion_config_block().get("syntax", {})
    
    def get_code_detection_config(self) -> Dict[str, Any]:
        """获取代码识别配置子表"""
        return self.get_conversion_config_block().get("code_detection", {})
    
    def get_horizontal_rule_config(self) -> Dict[str, Any]:
        """获取分隔符/分页符转换配置子表"""
        return self.get_conversion_config_block().get("horizontal_rule", {})
    
    # ========== 分隔符/分页符配置 ==========
    
    def is_horizontal_rule_enabled(self) -> bool:
        """
        是否启用分隔符转换
        
        返回:
            bool: 是否启用
        """
        config = self.get_horizontal_rule_config()
        enabled = config.get("enabled", True)
        safe_log.debug("分隔符转换启用: %s", enabled)
        return enabled
    
    def get_horizontal_rule_docx_to_md_config(self) -> Dict[str, str]:
        """
        获取文档转MD的分隔符配置
        
        返回:
            Dict[str, str]: {Word分隔符类型: MD分隔符}
                - page_break: 分页符对应的MD分隔符
                - section_break: 分节符对应的MD分隔符（所有类型统一）
                - horizontal_rule: 分隔线对应的MD分隔符（Horizontal Rule 1/2/3 样式统一转换）
        """
        config = self.get_horizontal_rule_config()
        docx_to_md = config.get("docx_to_md", {})
        defaults = {
            "page_break": "---",
            "section_break": "***",
            "horizontal_rule": "___"
        }
        result = {
            "page_break": docx_to_md.get("page_break", defaults["page_break"]),
            "section_break": docx_to_md.get("section_break", defaults["section_break"]),
            "horizontal_rule": docx_to_md.get("horizontal_rule", defaults["horizontal_rule"])
        }
        safe_log.debug("文档转MD分隔符配置: %s", result)
        return result
    
    def get_horizontal_rule_md_to_docx_config(self) -> Dict[str, str]:
        """
        获取MD转文档的分隔符配置
        
        返回:
            Dict[str, str]: {MD分隔符类型: Word分隔符}
                - dash: --- 对应的Word元素
                - asterisk: *** 对应的Word元素
                - underscore: ___ 对应的Word元素（可选 horizontal_rule_1/2/3）
        """
        config = self.get_horizontal_rule_config()
        md_to_docx = config.get("md_to_docx", {})
        defaults = {
            "dash": "page_break",
            "asterisk": "section_break",
            "underscore": "horizontal_rule_1"
        }
        result = {
            "dash": md_to_docx.get("dash", defaults["dash"]),
            "asterisk": md_to_docx.get("asterisk", defaults["asterisk"]),
            "underscore": md_to_docx.get("underscore", defaults["underscore"])
        }
        safe_log.debug("MD转文档分隔符配置: %s", result)
        return result
    
    def get_horizontal_rule_mapping(self, hr_type: str) -> str:
        """
        获取指定分隔符类型的转换目标（MD→DOCX）
        
        参数:
            hr_type: 分隔符类型 ("dash", "asterisk", "underscore")
            
        返回:
            str: 转换目标 ("page_break", "section_continuous", "section_break", "ignore")
        """
        config = self.get_horizontal_rule_md_to_docx_config()
        target = config.get(hr_type, "page_break")
        safe_log.debug("分隔符 %s 映射目标: %s", hr_type, target)
        return target
    
    def get_all_horizontal_rule_mappings(self) -> Dict[str, str]:
        """
        获取所有分隔符映射配置（MD→DOCX方向）
        
        返回:
            Dict[str, str]: {分隔符类型: 转换目标}
        """
        return self.get_horizontal_rule_md_to_docx_config()
    
    def get_md_separator_for_break_type(self, break_type: str) -> Optional[str]:
        """
        根据 Word 分页/分节/分隔线类型获取对应的 MD 分隔符（DOCX→MD）
        
        参数:
            break_type: Word 元素类型
                - "page_break": 分页符
                - "section_next", "section_continuous", "section_even", "section_odd": 分节符
                - "horizontal_rule": 分隔线（Horizontal Rule 1/2/3 样式）
            
        返回:
            str: MD 分隔符 ("---", "***", "___")，未找到或忽略时返回 None
        """
        config = self.get_horizontal_rule_docx_to_md_config()
        
        # 将所有分节符类型统一映射到 section_break 配置键
        type_to_config_key = {
            'page_break': 'page_break',
            'section_continuous': 'section_break',
            'section_next': 'section_break',
            'section_even': 'section_break',
            'section_odd': 'section_break',
            'horizontal_rule': 'horizontal_rule'
        }
        config_key = type_to_config_key.get(break_type, break_type)
        
        # 获取配置的MD分隔符
        md_separator = config.get(config_key)
        
        if md_separator and md_separator != "ignore":
            safe_log.debug("分隔符类型 %s (配置键: %s) 映射为: %s", break_type, config_key, md_separator)
            return md_separator
        
        safe_log.debug("分隔符类型 %s 映射为忽略或未配置", break_type)
        return None
    
    def get_docx_to_md_preserve_formatting(self) -> bool:
        """
        获取DOCX转MD时是否保留正文格式
        
        返回:
            bool: 是否保留格式（转为Markdown标记）
        """
        config = self.get_docx_to_md_config()
        preserve = config.get("preserve_formatting", True)
        safe_log.debug("DOCX转MD保留正文格式: %s", preserve)
        return preserve
    
    def get_docx_to_md_preserve_heading_formatting(self) -> bool:
        """
        获取DOCX转MD时是否保留小标题格式
        
        返回:
            bool: 是否保留小标题内的格式标记
        """
        config = self.get_docx_to_md_config()
        preserve = config.get("preserve_heading_formatting", False)
        safe_log.debug("DOCX转MD保留小标题格式: %s", preserve)
        return preserve
    
    def get_docx_to_md_preserve_table_header_formatting(self) -> bool:
        """
        获取DOCX转MD时是否保留表头格式
        
        返回:
            bool: 是否保留表头单元格中的格式标记
        """
        config = self.get_docx_to_md_config()
        preserve = config.get("preserve_table_header_formatting", False)
        safe_log.debug("DOCX转MD保留表头格式: %s", preserve)
        return preserve
    
    def get_md_to_docx_formatting_mode(self) -> str:
        """
        获取MD转DOCX正文格式处理模式
        
        返回:
            str: 处理模式 ("apply", "keep", "remove")
        """
        config = self.get_md_to_docx_config()
        mode = config.get("formatting_mode", "apply")
        if mode not in ["apply", "keep", "remove"]:
            mode = "apply"
        safe_log.debug("MD转DOCX正文格式处理模式: %s", mode)
        return mode
    
    def get_md_to_docx_heading_formatting_mode(self) -> str:
        """
        获取MD转DOCX小标题格式处理模式
        
        返回:
            str: 处理模式
                - "apply": 应用格式，覆盖样式默认格式（未标记文字显式不加粗/不斜体）
                - "keep": 保留标记原样
                - "remove": 清理标记，让Word标题样式格式自然生效
        """
        config = self.get_md_to_docx_config()
        mode = config.get("heading_formatting_mode", "remove")
        if mode not in ["apply", "keep", "remove"]:
            mode = "remove"
        safe_log.debug("MD转DOCX小标题格式处理模式: %s", mode)
        return mode
    
    def get_md_to_docx_table_header_formatting_mode(self) -> str:
        """
        获取MD转DOCX表头格式处理模式
        
        返回:
            str: 处理模式
                - "apply": 应用格式（**粗体** → 真正的粗体，可能与表格样式重复）
                - "keep": 保留标记原样
                - "remove": 清理标记，让表格样式（如 firstRow 加粗）自然生效
        """
        config = self.get_md_to_docx_config()
        mode = config.get("table_header_formatting_mode", "remove")
        if mode not in ["apply", "keep", "remove"]:
            mode = "remove"
        safe_log.debug("MD转DOCX表头格式处理模式: %s", mode)
        return mode
    
    def get_syntax_setting(self, key: str, default: str = None) -> str:
        """
        获取特定格式的Markdown语法配置
        
        参数:
            key: 格式键名（bold, italic, strikethrough, highlight, superscript, subscript）
            default: 默认值
            
        返回:
            str: 语法选项（如 "asterisk", "underscore", "extended", "html"）
        """
        config = self.get_formatting_syntax_config()
        value = config.get(key, default)
        safe_log.debug("获取语法配置 %s: %s", key, value)
        return value
    
    def get_all_syntax_settings(self) -> Dict[str, str]:
        """
        获取所有Markdown语法配置
        
        返回:
            Dict[str, str]: 语法配置字典
        """
        default_syntax = {
            "bold": "asterisk",
            "italic": "asterisk",
            "strikethrough": "extended",
            "highlight": "extended",
            "superscript": "html",
            "subscript": "html"
        }
        config = self.get_formatting_syntax_config()
        # 合并默认值和用户配置
        result = {**default_syntax, **config}
        safe_log.debug("获取所有语法配置: %s", result)
        return result
    
    def get_code_font(self) -> str:
        """
        获取代码字体
        
        返回:
            str: 字体名称
        """
        config = self.get_code_detection_config()
        font = config.get("code_font", "Consolas")
        return font
    
    def get_code_background_color(self) -> str:
        """
        获取代码背景颜色
        
        返回:
            str: 背景颜色（十六进制）
        """
        config = self.get_code_detection_config()
        color = config.get("code_background_color", "E7E6E6")
        return color
    
    def get_unordered_list_marker(self) -> str:
        """
        获取无序列表标记
        
        返回:
            str: 标记类型 ("dash", "asterisk", "plus")
        """
        config = self.get_formatting_syntax_config()
        marker = config.get("unordered_list", "dash")
        if marker not in ["dash", "asterisk", "plus"]:
            marker = "dash"
        safe_log.debug("获取无序列表标记: %s", marker)
        return marker
    
    def get_ordered_list_mode(self) -> str:
        """
        获取有序列表编号模式（暂未实现，预留）
        
        返回:
            str: 模式 ("restart", "preserve")
        """
        config = self.get_formatting_syntax_config()
        mode = config.get("ordered_list", "restart")
        if mode not in ["restart", "preserve"]:
            mode = "restart"
        safe_log.debug("获取有序列表模式: %s", mode)
        return mode
    
    def get_list_indent_spaces(self) -> int:
        """
        获取列表每级缩进的空格数
        
        用于 MD → DOCX 转换时识别多级列表嵌套级别。
        
        返回:
            int: 每级缩进的空格数，默认为 2
        """
        config = self.get_formatting_syntax_config()
        spaces = config.get("indent_spaces", 2)
        # 确保在有效范围内
        if not isinstance(spaces, int) or spaces < 1:
            spaces = 2
        safe_log.debug("获取列表缩进空格数: %d", spaces)
        return spaces
    
    def get_document_to_pdf_priority(self) -> List[str]:
        """获取文档转PDF的软件优先级列表"""
        special_conversions = self.get_special_conversions_config()
        return special_conversions.get("document_to_pdf", ["wps_writer", "msoffice_word", "libreoffice"])
    
    def get_spreadsheet_to_pdf_priority(self) -> List[str]:
        """获取表格转PDF的软件优先级列表"""
        special_conversions = self.get_special_conversions_config()
        return special_conversions.get("spreadsheet_to_pdf", ["wps_spreadsheets", "msoffice_excel", "libreoffice"])
    
    # --------------------------
    # 样式映射配置方法
    # --------------------------
    
    # ========== 代码样式配置 ==========
    
    def is_code_style_detection_enabled(self) -> bool:
        """
        是否启用代码样式检测（DOCX→MD）
        
        返回:
            bool: 是否启用代码样式检测（始终返回 True，因为样式列表控制具体行为）
        """
        # 新配置结构中没有 enabled 开关，始终返回 True
        return True
    
    def get_code_paragraph_styles(self) -> List[str]:
        """
        获取代码块段落样式列表（DOCX→MD）
        
        返回:
            List[str]: 段落样式名称列表，匹配后转为代码块 ```
        """
        config = self.get_style_code_block().get("docx_to_md", {})
        styles = config.get("paragraph_styles", [
            "HTML Preformatted", "Code Block", "代码块"
        ])
        safe_log.debug("代码块段落样式列表: %s", styles)
        return styles
    
    def get_code_character_styles(self) -> List[str]:
        """
        获取行内代码字符样式列表（DOCX→MD）
        
        返回:
            List[str]: 字符样式名称列表，匹配后转为行内代码 `
        """
        config = self.get_style_code_block().get("docx_to_md", {})
        styles = config.get("character_styles", [
            "HTML Code", "HTML Typewriter", "HTML Keyboard", "HTML Sample",
            "HTML Variable", "HTML Definition", "HTML Cite", "HTML Address",
            "HTML Acronym", "Inline Code", "Code", "Source Code",
            "行内代码", "代码", "源代码", "源码"
        ])
        safe_log.debug("行内代码字符样式列表: %s", styles)
        return styles
    
    def get_code_full_paragraph_as_block(self) -> bool:
        """
        获取是否将整段代码样式的段落视为代码块
        
        返回:
            bool: 是否整段转代码块
        """
        config = self.get_style_code_block().get("docx_to_md", {})
        return config.get("full_paragraph_as_block", True)
    
    def get_code_fuzzy_match_enabled(self) -> bool:
        """
        是否启用代码样式模糊匹配
        
        返回:
            bool: 是否启用模糊匹配
        """
        config = self.get_style_code_block().get("docx_to_md", {})
        enabled = config.get("fuzzy_match_enabled", True)
        safe_log.debug("代码样式模糊匹配启用: %s", enabled)
        return enabled
    
    def get_code_fuzzy_keywords(self) -> List[str]:
        """
        获取代码样式模糊匹配关键词（不区分大小写）
        
        返回:
            List[str]: 模糊匹配关键词列表
        """
        config = self.get_style_code_block().get("docx_to_md", {})
        keywords = config.get("fuzzy_keywords", ["code", "代码", "源码"])
        safe_log.debug("代码模糊匹配关键词: %s", keywords)
        return keywords
    
    # ========== 底纹检测配置 ==========
    
    def is_wps_shading_enabled(self) -> bool:
        """
        是否启用 WPS 底纹检测（DOCX→MD）
        
        WPS 底纹使用纯色填充方式（w:val="clear" + w:fill="D9D9D9"）
        
        返回:
            bool: 是否启用 WPS 底纹检测
        """
        config = self.get_style_code_block().get("docx_to_md", {})
        shading = config.get("shading", {})
        enabled = shading.get("wps_enabled", True)
        safe_log.debug("WPS 底纹检测启用: %s", enabled)
        return enabled
    
    def is_word_shading_enabled(self) -> bool:
        """
        是否启用 Word 底纹检测（DOCX→MD）
        
        Word 底纹使用百分比图案填充方式（w:val="pct15" + w:fill="FFFFFF"）
        
        返回:
            bool: 是否启用 Word 底纹检测
        """
        config = self.get_style_code_block().get("docx_to_md", {})
        shading = config.get("shading", {})
        enabled = shading.get("word_enabled", True)
        safe_log.debug("Word 底纹检测启用: %s", enabled)
        return enabled
    
    def get_inline_code_style(self) -> str:
        """
        获取行内代码样式配置（MD→DOCX）
        
        返回:
            str: 样式名
        """
        config = self.get_style_code_block().get("md_to_docx", {})
        style = config.get("inline_code_style", "Inline Code")
        safe_log.debug("行内代码样式: %s", style)
        return style
    
    def get_code_block_style(self) -> str:
        """
        获取代码块样式配置（MD→DOCX）
        
        返回:
            str: 样式名
        """
        config = self.get_style_code_block().get("md_to_docx", {})
        style = config.get("code_block_style", "Code Block")
        safe_log.debug("代码块样式: %s", style)
        return style
    
    # ========== 引用样式配置 ==========
    
    def is_quote_style_detection_enabled(self) -> bool:
        """
        是否启用引用样式检测（DOCX→MD）
        
        返回:
            bool: 是否启用引用样式检测（始终返回 True，因为样式列表控制具体行为）
        """
        # 新配置结构中没有 enabled 开关，始终返回 True
        return True
    
    def get_quote_level_styles(self) -> Dict[str, int]:
        """
        获取分级引用样式映射（DOCX→MD）
        
        返回:
            Dict[str, int]: 样式名到引用级别的映射 {"quote 1": 1, "引用 1": 1, ...}
        """
        config = self.get_style_quote_block().get("docx_to_md", {})
        level_styles = config.get("level_styles", {
            "quote 1": 1, "quote 2": 2, "quote 3": 3, "quote 4": 4, "quote 5": 5,
            "quote 6": 6, "quote 7": 7, "quote 8": 8, "quote 9": 9,
            "引用 1": 1, "引用 2": 2, "引用 3": 3, "引用 4": 4, "引用 5": 5,
            "引用 6": 6, "引用 7": 7, "引用 8": 8, "引用 9": 9
        })
        safe_log.debug("分级引用样式映射: %d 个样式", len(level_styles))
        return level_styles
    
    def get_quote_paragraph_styles(self) -> List[str]:
        """
        获取通用引用段落样式列表（DOCX→MD）
        
        返回:
            List[str]: 通用引用段落样式列表（无法确定级别时默认为1级）
        """
        config = self.get_style_quote_block().get("docx_to_md", {})
        styles = config.get("paragraph_styles", [
            "Quote", "Block Text", "Intense Quote", "引用", "明显引用"
        ])
        safe_log.debug("通用引用段落样式列表: %s", styles)
        return styles
    
    def get_quote_character_styles(self) -> List[str]:
        """
        获取引用字符样式列表（DOCX→MD）
        
        返回:
            List[str]: 引用字符样式列表（转为行内代码）
        """
        config = self.get_style_quote_block().get("docx_to_md", {})
        styles = config.get("character_styles", ["Quote Char", "引用字符"])
        safe_log.debug("引用字符样式列表: %s", styles)
        return styles
    
    def get_quote_full_paragraph_as_block(self) -> bool:
        """
        获取是否将整段引用样式的段落视为引用块
        
        返回:
            bool: 是否整段转引用块
        """
        config = self.get_style_quote_block().get("docx_to_md", {})
        return config.get("full_paragraph_as_block", True)
    
    def get_quote_fuzzy_match_enabled(self) -> bool:
        """
        是否启用引用样式模糊匹配
        
        返回:
            bool: 是否启用模糊匹配
        """
        config = self.get_style_quote_block().get("docx_to_md", {})
        enabled = config.get("fuzzy_match_enabled", True)
        safe_log.debug("引用样式模糊匹配启用: %s", enabled)
        return enabled
    
    def get_quote_fuzzy_keywords(self) -> List[str]:
        """
        获取引用样式模糊匹配关键词（不区分大小写）
        
        返回:
            List[str]: 模糊匹配关键词列表
        """
        config = self.get_style_quote_block().get("docx_to_md", {})
        keywords = config.get("fuzzy_keywords", ["quote", "引用"])
        safe_log.debug("引用模糊匹配关键词: %s", keywords)
        return keywords
    
    def get_quote_style_for_level(self, level: int) -> str:
        """
        根据引用级别获取样式配置（MD→DOCX）
        
        参数:
            level: 引用级别 (1-9)
            
        返回:
            str: 样式名，如 "Quote 1"
        """
        # 确保级别在有效范围内
        level = max(1, min(9, level))
        
        config = self.get_style_quote_block().get("md_to_docx", {})
        
        style_key = f"level_{level}_style"
        style = config.get(style_key, f"Quote {level}")
        
        safe_log.debug("引用级别 %d 样式: %s", level, style)
        return style
    
    # ========== 公式样式配置 ==========
    
    def get_inline_formula_style(self) -> str:
        """
        获取行内公式样式配置（MD→DOCX）
        
        返回:
            str: 样式名
        """
        config = self.get_style_formula_block().get("md_to_docx", {})
        style = config.get("inline_formula_style", "Inline Formula")
        safe_log.debug("行内公式样式: %s", style)
        return style
    
    def get_formula_block_style(self) -> str:
        """
        获取公式块样式配置（MD→DOCX）
        
        返回:
            str: 样式名
        """
        config = self.get_style_formula_block().get("md_to_docx", {})
        style = config.get("formula_block_style", "Formula Block")
        safe_log.debug("公式块样式: %s", style)
        return style
    
    # ========== 表格样式配置 ==========
    
    def get_table_style(self) -> str:
        """
        获取表格样式配置（MD→DOCX）
        
        返回:
            str: 样式名
        """
        config = self.get_style_table_block().get("md_to_docx", {})
        style = config.get("table_style", "Three Line Table")
        safe_log.debug("表格样式: %s", style)
        return style
    
    def get_table_content_style(self) -> str:
        """
        获取表格内容样式配置（MD→DOCX）
        
        返回:
            str: 样式名
        """
        config = self.get_style_table_block().get("md_to_docx", {})
        style = config.get("table_content_style", "Table Content")
        safe_log.debug("表格内容样式: %s", style)
        return style
    
    def get_table_style_mode(self) -> str:
        """
        获取表格样式模式（MD→DOCX）
        
        返回:
            str: 模式 ("builtin" 使用内置样式, "custom" 使用自定义样式名)
        """
        config = self.get_style_table_block().get("md_to_docx", {})
        mode = config.get("table_style_mode", "builtin")
        if mode not in ["builtin", "custom"]:
            mode = "builtin"
        safe_log.debug("表格样式模式: %s", mode)
        return mode
    
    def get_builtin_table_style_key(self) -> str:
        """
        获取内置表格样式键名（MD→DOCX）
        
        返回:
            str: 内置样式键名 ("three_line_table" 或 "table_grid")
        """
        config = self.get_style_table_block().get("md_to_docx", {})
        key = config.get("builtin_style_key", "three_line_table")
        if key not in ["three_line_table", "table_grid"]:
            key = "three_line_table"
        safe_log.debug("内置表格样式键名: %s", key)
        return key
    
    def get_custom_table_style_name(self) -> str:
        """
        获取自定义表格样式名称（MD→DOCX）
        
        返回:
            str: 用户自定义的样式名称
        """
        config = self.get_style_table_block().get("md_to_docx", {})
        name = config.get("custom_style_name", "")
        safe_log.debug("自定义表格样式名: %s", name)
        return name
    
    def reload_configs(self):
        """重新加载所有配置文件（无日志依赖）"""
        safe_log.info("重新加载所有配置文件...")
        self._initialized = False
        self._configs = {}
        self._initialize(self._config_dir)
    
    # --------------------------
    # 第四层：配置修改方法
    # --------------------------
    
    def update_config_value(self, config_name: str, section: str, key: str, value: Any) -> bool:
        """
        更新配置文件中的特定值
        
        参数:
            config_name: 配置名称（如"gui_config", "typo_settings"等）
            section: 节名称（可以是多级，如"window.size"）
            key: 键名称
            value: 新值
            
        返回:
            bool: 更新是否成功
        """
        if config_name not in CONFIG_FILES:
            safe_log.error("未知的配置名称: %s", config_name)
            return False
        
        filename = CONFIG_FILES[config_name]
        filepath = os.path.join(self._config_dir, filename)
        
        success = update_toml_value(filepath, section, key, value)
        if success:
            # 重新加载该配置文件以更新内存中的配置
            self._configs[config_name] = self._load_single_config(filename)
            safe_log.info("配置更新成功: %s -> %s.%s = %s", config_name, section, key, value)
        
        return success
    
    def update_config_section(self, config_name: str, section: str, data: Dict[str, Any]) -> bool:
        """
        更新配置文件的整个节（保留注释和原有顺序）
        
        参数:
            config_name: 配置名称（如"gui_config", "typo_settings"等）
            section: 节名称（可以是多级，如"window.size"）
            data: 新的节数据
            
        返回:
            bool: 更新是否成功
        """
        if config_name not in CONFIG_FILES:
            safe_log.error("未知的配置名称: %s", config_name)
            return False
        
        filename = CONFIG_FILES[config_name]
        filepath = os.path.join(self._config_dir, filename)
        
        try:
            # 读取现有配置（保留注释）
            from .toml_operations import read_toml_document, write_toml_document
            doc = read_toml_document(filepath)
            if doc is None:
                # 如果文件不存在或读取失败，创建新文档
                from tomlkit import document
                doc = document()
            
            # 分割多级节名称
            section_parts = section.split('.')
            
            # 导航到目标节
            current = doc
            for part in section_parts:
                if part not in current:
                    # 创建不存在的节
                    from tomlkit import table
                    current[part] = table()
                    current = current[part]
                else:
                    current = current[part]
            
            # 更新整个节 - 保留注释和原有顺序
            # 1. 获取原有的所有键（保持顺序）
            existing_keys = list(current.keys())
            
            # 2. 先更新已存在的键（保留注释，保持原有顺序）
            for key in existing_keys:
                if key in data:
                    # 键存在于新数据中，更新其值（保留注释）
                    current[key] = data[key]
                else:
                    # 键不存在于新数据中，删除它
                    del current[key]
            
            # 3. 再添加新键（原来不存在的键添加到末尾）
            for key in data:
                if key not in existing_keys:
                    current[key] = data[key]
            
            # 写回文件（保留注释）
            success = write_toml_document(filepath, doc)
            if success:
                # 重新加载该配置文件以更新内存中的配置
                self._configs[config_name] = self._load_single_config(filename)
                safe_log.info("配置节更新成功: %s -> %s", config_name, section)
            
            return success
            
        except Exception as e:
            safe_log.error("更新配置节失败: %s -> %s | 错误: %s", config_name, section, str(e))
            return False

# 创建全局配置管理器实例
config_manager = ConfigManager()
