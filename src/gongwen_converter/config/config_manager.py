"""
配置管理单例类
提供配置的加载、访问和修改功能，支持GUI配置
"""

import os
from typing import Dict, Any, List, Tuple, Optional, Union
from .safe_logger import safe_log
from .constants import DEFAULT_CONFIG, CONFIG_FILES
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
    # 第一层：配置块级别访问方法
    # --------------------------
    def get_logger_config_block(self) -> Dict[str, Any]:
        """获取整个日志配置块"""
        return self._configs.get("logger_config", {})
    
    def get_symbol_settings_block(self) -> Dict[str, Any]:
        """获取整个符号校对配置块"""
        return self._configs.get("symbol_settings", {})

    def get_typos_settings_block(self) -> Dict[str, Any]:
        """获取整个错别字配置块"""
        return self._configs.get("typos_settings", {})

    def get_sensitive_words_settings_block(self) -> Dict[str, Any]:
        """获取整个敏感词配置块"""
        return self._configs.get("sensitive_words_settings", {})
    
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
    
    def get_image_config_block(self) -> Dict[str, Any]:
        """获取整个图片配置块"""
        return self._configs.get("image_config", {})
    
    # --------------------------
    # 第二层：子表级别访问方法
    # --------------------------
    
    # 日志配置子表
    def get_logging_config(self) -> Dict[str, Any]:
        """获取日志配置子表"""
        return self.get_logger_config_block().get("logging", {})
    
    # 错别字/符号等校对配置子表
    def get_symbol_engine_settings(self) -> Dict[str, Any]:
        """获取符号校对引擎设置子表"""
        return self.get_symbol_settings_block().get("engine_settings", {})

    def get_typos_engine_settings(self) -> Dict[str, Any]:
        """获取错别字引擎设置子表"""
        return self.get_typos_settings_block().get("engine_settings", {})

    def get_sensitive_words_engine_settings(self) -> Dict[str, Any]:
        """获取敏感词引擎设置子表"""
        return self.get_sensitive_words_settings_block().get("engine_settings", {})

    def get_symbol_pairing_config(self) -> Dict[str, Any]:
        """获取符号配对子表"""
        return self.get_symbol_settings_block().get("symbol_pairing", {})
    
    def get_symbol_map_config(self) -> Dict[str, Any]:
        """获取符号映射子表"""
        return self.get_symbol_settings_block().get("symbol_map", {})
    
    def get_typos_config(self) -> Dict[str, Any]:
        """获取错别字子表"""
        return self.get_typos_settings_block().get("typos", {})

    def get_sensitive_words_config(self) -> Dict[str, Any]:
        """获取敏感词子表"""
        return self.get_sensitive_words_settings_block().get("sensitive_words", {})
    
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
    
    def get_extraction_defaults_config(self) -> Dict[str, Any]:
        """获取图片提取默认设置子表"""
        return self.get_image_config_block().get("extraction_defaults", {})

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
        """获取标点符号映射（带默认值）"""
        symbol_map = self.get_symbol_map_config()
        return symbol_map if isinstance(symbol_map, dict) else {}
    
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
    
    def get_docx_to_md_keep_images(self) -> bool:
        """
        获取文档转MD时是否保留图片
        
        返回:
            bool: 是否保留图片
        """
        extraction_config = self.get_extraction_defaults_config()
        keep_images = extraction_config.get("docx_to_md_keep_images", True)
        safe_log.debug("文档转MD保留图片: %s", keep_images)
        return keep_images
    
    def get_docx_to_md_enable_ocr(self) -> bool:
        """
        获取文档转MD时是否启用OCR
        
        返回:
            bool: 是否启用OCR
        """
        extraction_config = self.get_extraction_defaults_config()
        enable_ocr = extraction_config.get("docx_to_md_enable_ocr", False)
        safe_log.debug("文档转MD启用OCR: %s", enable_ocr)
        return enable_ocr
    
    def get_xlsx_to_md_keep_images(self) -> bool:
        """
        获取表格转MD时是否保留图片
        
        返回:
            bool: 是否保留图片
        """
        extraction_config = self.get_extraction_defaults_config()
        keep_images = extraction_config.get("xlsx_to_md_keep_images", True)
        safe_log.debug("表格转MD保留图片: %s", keep_images)
        return keep_images
    
    def get_xlsx_to_md_enable_ocr(self) -> bool:
        """
        获取表格转MD时是否启用OCR
        
        返回:
            bool: 是否启用OCR
        """
        extraction_config = self.get_extraction_defaults_config()
        enable_ocr = extraction_config.get("xlsx_to_md_enable_ocr", False)
        safe_log.debug("表格转MD启用OCR: %s", enable_ocr)
        return enable_ocr
    
    def get_layout_to_md_keep_images(self) -> bool:
        """
        获取版式文件转MD时是否保留图片
        
        返回:
            bool: 是否保留图片
        """
        extraction_config = self.get_extraction_defaults_config()
        keep_images = extraction_config.get("layout_to_md_keep_images", True)
        safe_log.debug("版式转MD保留图片: %s", keep_images)
        return keep_images
    
    def get_layout_to_md_enable_ocr(self) -> bool:
        """
        获取版式文件转MD时是否启用OCR
        
        返回:
            bool: 是否启用OCR
        """
        extraction_config = self.get_extraction_defaults_config()
        enable_ocr = extraction_config.get("layout_to_md_enable_ocr", False)
        safe_log.debug("版式转MD启用OCR: %s", enable_ocr)
        return enable_ocr
    
    def get_image_to_md_keep_images(self) -> bool:
        """
        获取图片文件转MD时是否保留图片
        
        返回:
            bool: 是否保留图片
        """
        extraction_config = self.get_extraction_defaults_config()
        keep_images = extraction_config.get("image_to_md_keep_images", True)
        safe_log.debug("图片转MD保留图片: %s", keep_images)
        return keep_images
    
    def get_image_to_md_enable_ocr(self) -> bool:
        """
        获取图片文件转MD时是否启用OCR
        
        返回:
            bool: 是否启用OCR
        """
        extraction_config = self.get_extraction_defaults_config()
        enable_ocr = extraction_config.get("image_to_md_enable_ocr", True)
        safe_log.debug("图片转MD启用OCR: %s", enable_ocr)
        return enable_ocr
    
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
                - image_link_format: str - 图片链接格式 ("markdown" 或 "wiki")
                - image_embed: bool - 图片是否嵌入显示
                - md_file_link_format: str - MD文件链接格式 ("markdown" 或 "wiki")
                - md_file_embed: bool - MD文件是否嵌入显示（仅wiki有效）
        """
        link_format_config = self.get_link_format_config()
        settings = {
            "image_link_format": link_format_config.get("image_link_format", "wiki"),
            "image_embed": link_format_config.get("image_embed", True),
            "md_file_link_format": link_format_config.get("md_file_link_format", "wiki"),
            "md_file_embed": link_format_config.get("md_file_embed", True)
        }
        safe_log.debug("获取Markdown链接格式设置: %s", settings)
        return settings
    
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
    def is_embedding_enabled(self) -> bool:
        """
        是否启用嵌入功能
        
        返回:
            bool: 是否启用嵌入功能总开关
        """
        config = self.get_embed_links_config()
        enabled = config.get("enabled", True)
        safe_log.debug("嵌入功能启用状态: %s", enabled)
        return enabled
    
    def get_embed_image_mode(self) -> str:
        """
        获取嵌入图片处理模式
        
        返回:
            str: 处理模式 ("keep", "extract_text", "remove", "embed")
        """
        config = self.get_embed_links_config()
        mode = config.get("image_mode", "embed")
        safe_log.debug("获取嵌入图片处理模式: %s", mode)
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
    
    def get_document_to_pdf_priority(self) -> List[str]:
        """获取文档转PDF的软件优先级列表"""
        special_conversions = self.get_special_conversions_config()
        return special_conversions.get("document_to_pdf", ["wps_writer", "msoffice_word", "libreoffice"])
    
    def get_spreadsheet_to_pdf_priority(self) -> List[str]:
        """获取表格转PDF的软件优先级列表"""
        special_conversions = self.get_special_conversions_config()
        return special_conversions.get("spreadsheet_to_pdf", ["wps_spreadsheets", "msoffice_excel", "libreoffice"])

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
