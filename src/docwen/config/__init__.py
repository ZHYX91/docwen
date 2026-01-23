"""
配置读取包

提供配置管理功能，包括：
    - config_manager: 配置管理器单例
    - SOFTWARE_ID_MAPPING: 软件标识符映射
    - DEFAULT_CONFIG: 默认配置合集
    - CONFIG_FILES: 配置文件映射

使用方式：
    from docwen.config import config_manager
    
    # 获取配置
    theme = config_manager.get_default_theme()
    
    # 导入常量
    from docwen.config import SOFTWARE_ID_MAPPING
"""

from .config_manager import config_manager
from .schemas import (
    SOFTWARE_ID_MAPPING,
    DEFAULT_CONFIG,
    CONFIG_FILES,
)

__all__ = [
    "config_manager",
    "SOFTWARE_ID_MAPPING",
    "DEFAULT_CONFIG",
    "CONFIG_FILES",
]
