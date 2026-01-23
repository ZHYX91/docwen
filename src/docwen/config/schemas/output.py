"""
输出配置模块

对应配置文件：output_config.toml

包含：
    - DEFAULT_OUTPUT_CONFIG: 默认输出配置
    - OutputConfigMixin: 输出配置获取方法
"""

from typing import Dict, Any
from ..safe_logger import safe_log

# ==============================================================================
#                              默认配置
# ==============================================================================

DEFAULT_OUTPUT_CONFIG = {
    "output_config": {
        "directory": {
            "mode": "source",              # 输出目录模式: source/custom
            "custom_path": "",             # 自定义输出路径
            "create_date_subfolder": False,  # 是否创建日期子目录
            "date_folder_format": "%Y-%m-%d"  # 日期子目录格式
        },
        "behavior": {
            "auto_open_folder": False      # 转换完成后是否自动打开文件夹
        },
        "intermediate_files": {
            "save_to_output": True         # 是否保存中间文件到输出目录
        }
    }
}

# 配置文件名
CONFIG_FILE = "output_config.toml"

# ==============================================================================
#                              Mixin 类
# ==============================================================================

class OutputConfigMixin:
    """
    输出配置获取方法 Mixin
    
    提供输出目录和行为相关配置的访问方法。
    """
    
    # --------------------------------------------------------------------------
    # 第一层：配置块
    # --------------------------------------------------------------------------
    
    def get_output_config_block(self) -> Dict[str, Any]:
        """
        获取整个输出配置块
        
        返回:
            Dict[str, Any]: 输出配置字典，包含 directory、behavior、intermediate_files 子表
        """
        return self._configs.get("output_config", {})
    
    # --------------------------------------------------------------------------
    # 第二层：子表
    # --------------------------------------------------------------------------
    
    def get_output_directory_settings(self) -> Dict[str, Any]:
        """
        获取输出目录设置子表
        
        返回:
            Dict[str, Any]: 目录设置字典
        """
        return self.get_output_config_block().get("directory", {})
    
    def get_output_behavior_settings(self) -> Dict[str, Any]:
        """
        获取输出行为设置子表
        
        返回:
            Dict[str, Any]: 行为设置字典
        """
        return self.get_output_config_block().get("behavior", {})
    
    def get_output_intermediate_files_config(self) -> Dict[str, Any]:
        """
        获取输出中间文件配置子表
        
        返回:
            Dict[str, Any]: 中间文件配置字典
        """
        return self.get_output_config_block().get("intermediate_files", {})
    
    # --------------------------------------------------------------------------
    # 第三层：具体配置值
    # --------------------------------------------------------------------------
    
    def get_output_directory_mode(self) -> str:
        """
        获取输出目录模式
        
        返回:
            str: 输出目录模式 ("source"=源文件目录, "custom"=自定义目录)
        """
        directory_settings = self.get_output_directory_settings()
        mode = directory_settings.get("mode", "source")
        safe_log.debug("获取输出目录模式: %s", mode)
        return mode
    
    def get_custom_output_path(self) -> str:
        """
        获取自定义输出路径
        
        返回:
            str: 自定义输出路径，仅当 mode="custom" 时有效
        """
        directory_settings = self.get_output_directory_settings()
        path = directory_settings.get("custom_path", "")
        safe_log.debug("获取自定义输出路径: %s", path)
        return path
    
    def get_create_date_subfolder(self) -> bool:
        """
        获取是否创建日期子目录
        
        返回:
            bool: 是否创建日期子目录
        """
        directory_settings = self.get_output_directory_settings()
        create = directory_settings.get("create_date_subfolder", False)
        safe_log.debug("获取创建日期子目录: %s", create)
        return create
    
    def get_date_folder_format(self) -> str:
        """
        获取日期子目录格式
        
        返回:
            str: 日期格式字符串（如 "%Y-%m-%d"）
        """
        directory_settings = self.get_output_directory_settings()
        fmt = directory_settings.get("date_folder_format", "%Y-%m-%d")
        safe_log.debug("获取日期目录格式: %s", fmt)
        return fmt
    
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
    
    def get_save_intermediate_files(self) -> bool:
        """
        获取是否保存中间文件到输出目录
        
        返回:
            bool: 是否保存中间文件
        """
        intermediate_config = self.get_output_intermediate_files_config()
        save_to_output = intermediate_config.get("save_to_output", True)
        safe_log.debug("保存中间文件设置: %s", save_to_output)
        return save_to_output
