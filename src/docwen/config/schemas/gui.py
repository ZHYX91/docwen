"""
GUI 配置模块

对应配置文件：gui_config.toml

详细说明：
    包含 GUI 界面的默认配置和访问方法。
    GUI 配置控制窗口布局、主题、透明度、语言等界面相关设置。

包含：
    - DEFAULT_GUI_CONFIG: 默认 GUI 配置
    - GUIConfigMixin: GUI 配置获取方法

依赖：
    - safe_logger: 安全日志记录（用于配置访问时的调试日志）

使用方式：
    # 通过 ConfigManager 访问（推荐）
    from docwen.config import config_manager
    
    theme = config_manager.get_default_theme()
    width = config_manager.get_center_panel_width()
    
    # 直接导入默认配置
    from docwen.config.schemas.gui import DEFAULT_GUI_CONFIG
"""

from typing import Dict, Any, Tuple

from ..safe_logger import safe_log


# ==============================================================================
#                              默认配置
# ==============================================================================

DEFAULT_GUI_CONFIG = {
    "gui_config": {
        "window": {
            "center_panel_width": 400,
            "left_panel_width": 400,          # 批量面板宽度
            "right_panel_width": 300,         # 模板面板宽度
            "center_panel_screen_x": 0,
            "window_y": 0,
            "default_mode": "single",
            "default_height": 740,
            "min_height": 720,
            "auto_center": True,
            "remember_gui_state": True
        },
        "component": {
            "file_drop_height": 200
        },
        "theme": {
            "default_theme": "morph"
        },
        "transparency": {
            "enabled": True,
            "default_value": 0.95,
            "min_value": 0.8,
            "max_value": 1.0
        },
        "template": {
            "md_default_template": "docx"
        },
        "language": {
            "locale": "zh_CN"  # 默认语言: zh_CN（简体中文）, en_US（英文）
        }
    }
}

# 配置文件名
CONFIG_FILE = "gui_config.toml"


# ==============================================================================
#                              Mixin 类
# ==============================================================================

class GUIConfigMixin:
    """
    GUI 配置获取方法 Mixin
    
    提供 GUI 相关配置的访问方法，包括窗口设置、主题、透明度等。
    
    注意：
        此类设计为 Mixin，需要与 ConfigManager 一起使用。
        假定宿主类具有 _configs 属性（配置字典）。
    
    配置结构：
        gui_config:
            window: 窗口布局设置
            component: 组件设置
            theme: 主题设置
            transparency: 透明度设置
            template: 模板设置
            language: 语言设置
    """
    
    # 类型提示：声明 _configs 属性（由 ConfigManager 提供）
    _configs: Dict[str, Dict[str, Any]]
    
    # --------------------------------------------------------------------------
    # 第一层：配置块
    # --------------------------------------------------------------------------
    
    def get_gui_config_block(self) -> Dict[str, Any]:
        """
        获取整个 GUI 配置块
        
        返回：
            Dict[str, Any]: GUI 配置字典，包含 window、theme、transparency 等子表
        
        示例：
            {
                "window": {...},
                "component": {...},
                "theme": {...},
                "transparency": {...},
                "template": {...},
                "language": {...}
            }
        """
        return self._configs.get("gui_config", {})
    
    # --------------------------------------------------------------------------
    # 第二层：子表
    # --------------------------------------------------------------------------
    
    def get_window_config(self) -> Dict[str, Any]:
        """
        获取窗口设置子表
        
        返回：
            Dict[str, Any]: 窗口设置字典，包含尺寸、位置、模式等
        """
        return self.get_gui_config_block().get("window", {})

    def get_component_config(self) -> Dict[str, Any]:
        """
        获取组件设置子表
        
        返回：
            Dict[str, Any]: 组件设置字典
        """
        return self.get_gui_config_block().get("component", {})
    
    def get_theme_config(self) -> Dict[str, Any]:
        """
        获取主题设置子表
        
        返回：
            Dict[str, Any]: 主题设置字典
        """
        return self.get_gui_config_block().get("theme", {})
    
    def get_transparency_config(self) -> Dict[str, Any]:
        """
        获取透明度设置子表
        
        返回：
            Dict[str, Any]: 透明度设置字典
        """
        return self.get_gui_config_block().get("transparency", {})

    def get_template_config(self) -> Dict[str, Any]:
        """
        获取模板设置子表（从 GUI 配置）
        
        返回：
            Dict[str, Any]: 模板设置字典
        """
        return self.get_gui_config_block().get("template", {})
    
    def get_language_config(self) -> Dict[str, Any]:
        """
        获取语言设置子表（从 GUI 配置）
        
        返回：
            Dict[str, Any]: 语言设置字典
        """
        return self.get_gui_config_block().get("language", {})
    
    # --------------------------------------------------------------------------
    # 第三层：窗口配置具体值
    # --------------------------------------------------------------------------
    
    def get_center_panel_width(self) -> int:
        """
        获取中栏宽度
        
        返回：
            int: 中栏宽度（像素）
        """
        window_config = self.get_window_config()
        width = window_config.get("center_panel_width", 400)
        safe_log.debug("获取中栏宽度: %d", width)
        return width
    
    def get_batch_panel_width(self) -> int:
        """
        获取批量面板宽度

        返回：
            int: 批量面板宽度（像素）
        """
        window_config = self.get_window_config()
        width = window_config.get("left_panel_width", 400)
        safe_log.debug("获取批量面板宽度: %d", width)
        return width

    def get_template_panel_width(self) -> int:
        """
        获取模板面板宽度

        返回：
            int: 模板面板宽度（像素）
        """
        window_config = self.get_window_config()
        width = window_config.get("right_panel_width", 300)
        safe_log.debug("获取模板面板宽度: %d", width)
        return width
    
    def get_center_panel_screen_x(self) -> int:
        """
        获取中栏在屏幕上的 X 坐标
        
        返回：
            int: 中栏屏幕 X 坐标
        """
        window_config = self.get_window_config()
        x = window_config.get("center_panel_screen_x", 0)
        safe_log.debug("获取中栏屏幕X坐标: %d", x)
        return x
    
    def get_window_y(self) -> int:
        """
        获取窗口 Y 坐标
        
        返回：
            int: 窗口 Y 坐标
        """
        window_config = self.get_window_config()
        y = window_config.get("window_y", 0)
        safe_log.debug("获取窗口Y坐标: %d", y)
        return y
    
    def get_default_mode(self) -> str:
        """
        获取默认启动模式
        
        返回：
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
        
        返回：
            int: 窗口高度（像素）
        """
        window_config = self.get_window_config()
        height = window_config.get("default_height", 740)
        safe_log.debug("获取窗口高度: %d", height)
        return height

    def get_min_height(self) -> int:
        """
        获取窗口最小高度
        
        返回：
            int: 最小高度（像素）
        """
        window_config = self.get_window_config()
        min_height = window_config.get("min_height", 720)
        safe_log.debug("获取窗口最小高度: %d", min_height)
        return min_height

    def get_window_position(self) -> Tuple[int, int]:
        """
        获取窗口默认位置（基于中栏屏幕坐标计算）
        
        返回：
            Tuple[int, int]: (x 坐标, y 坐标)
        """
        window_config = self.get_window_config()
        
        # 获取中栏屏幕坐标和窗口 Y 坐标
        center_x = window_config.get("center_panel_screen_x", 0)
        y = window_config.get("window_y", 0)
        
        # 单文件模式：窗口 X = 中栏 X（中栏居中显示）
        window_x = center_x
        
        safe_log.debug("获取窗口位置: (%d, %d) [从中栏坐标%d计算]", window_x, y, center_x)
        return window_x, y
    
    def should_remember_gui_state(self) -> bool:
        """
        检查是否应记住窗口位置
        
        返回：
            bool: 是否记住窗口位置
        """
        window_config = self.get_window_config()
        remember = window_config.get("remember_gui_state", True)
        safe_log.debug("记住窗口位置: %s", remember)
        return remember

    def should_auto_center(self) -> bool:
        """
        检查是否应自动居中窗口
        
        返回：
            bool: 是否自动居中窗口
        """
        window_config = self.get_window_config()
        auto_center = window_config.get("auto_center", True)
        safe_log.debug("自动居中窗口: %s", auto_center)
        return auto_center
    
    # --------------------------------------------------------------------------
    # 第三层：组件配置具体值
    # --------------------------------------------------------------------------
    
    def get_file_drop_height(self) -> int:
        """
        获取文件拖拽区域高度
        
        返回：
            int: 文件拖拽区域高度（像素）
        """
        component_config = self.get_component_config()
        height = component_config.get("file_drop_height", 200)
        safe_log.debug("获取文件拖拽区域高度: %d", height)
        return height
    
    # --------------------------------------------------------------------------
    # 第三层：主题配置具体值
    # --------------------------------------------------------------------------

    def get_default_theme(self) -> str:
        """
        获取默认主题名称
        
        返回：
            str: 主题名称
        """
        theme_config = self.get_theme_config()
        theme = theme_config.get("default_theme", "morph")
        safe_log.debug("获取默认主题: %s", theme)
        return theme
    
    # --------------------------------------------------------------------------
    # 第三层：透明度配置具体值
    # --------------------------------------------------------------------------

    def is_transparency_enabled(self) -> bool:
        """
        检查是否启用透明度效果
        
        返回：
            bool: 是否启用透明度
        """
        transparency_config = self.get_transparency_config()
        enabled = transparency_config.get("enabled", True)
        safe_log.debug("透明度启用状态: %s", enabled)
        return enabled

    def get_transparency_value(self) -> float:
        """
        获取透明度值
        
        返回：
            float: 透明度值 (0.0-1.0)
        """
        transparency_config = self.get_transparency_config()
        transparency = transparency_config.get("default_value", 0.95)
        # 确保在有效范围内
        transparency = max(0.1, min(1.0, transparency))
        safe_log.debug("获取透明度值: %.2f", transparency)
        return transparency
    
    # --------------------------------------------------------------------------
    # 第三层：模板配置具体值
    # --------------------------------------------------------------------------

    def get_default_md_template_type(self) -> str:
        """
        获取默认 MD 文件模板类型（从 GUI 配置）
        
        返回：
            str: 默认模板类型 ("docx" 或 "xlsx")
        """
        template_config = self.get_template_config()
        template_type = template_config.get("md_default_template", "docx")
        # 确保返回有效值
        if template_type not in ["docx", "xlsx"]:
            template_type = "docx"
        safe_log.debug("获取默认MD模板类型: %s", template_type)
        return template_type
    
    # --------------------------------------------------------------------------
    # 第三层：语言配置具体值
    # --------------------------------------------------------------------------
    
    def get_locale(self) -> str:
        """
        获取当前语言设置（从 GUI 配置）
        
        返回：
            str: 语言代码，如 "zh_CN" 或 "en_US"
        """
        language_config = self.get_language_config()
        locale = language_config.get("locale", "zh_CN")
        safe_log.debug("获取语言设置: %s", locale)
        return locale
