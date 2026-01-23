"""
配置管理单例类

提供配置的加载、访问和修改功能，支持 GUI 配置和 TOML 文件持久化。

架构：
    - 核心类负责：单例管理、配置加载、配置合并、配置修改
    - Mixin 类提供：各功能域的 getter 方法

依赖：
    - safe_logger: 安全日志记录
    - schemas: 配置模块（DEFAULT_CONFIG, CONFIG_FILES, 各 Mixin）
    - toml_operations: TOML 文件读写

使用方式：
    from docwen.config import config_manager
    
    theme = config_manager.get_default_theme()
    log_level = config_manager.get_log_level()
"""

import os
from typing import Dict, Any

from .safe_logger import safe_log
from .schemas import (
    DEFAULT_CONFIG,
    CONFIG_FILES,
    LoggerConfigMixin,
    GUIConfigMixin,
    ProofreadConfigMixin,
    OutputConfigMixin,
    SoftwareConfigMixin,
    LinkConfigMixin,
    NumberingConfigMixin,
    StyleConfigMixin,
    ConversionConfigMixin,
    OptimizationConfigMixin,
)
from .toml_operations import read_toml_file, update_toml_value


class ConfigManager(
    LoggerConfigMixin,
    GUIConfigMixin,
    ProofreadConfigMixin,
    OutputConfigMixin,
    SoftwareConfigMixin,
    LinkConfigMixin,
    NumberingConfigMixin,
    StyleConfigMixin,
    ConversionConfigMixin,
    OptimizationConfigMixin,
):
    """
    配置管理单例类
    
    通过 Mixin 模式组合各功能域的配置获取方法。
    核心类负责配置的加载、合并和修改。
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
        """初始化配置管理器"""
        if self._initialized:
            return
            
        # 使用路径工具获取配置目录
        try:
            from docwen.utils.path_utils import get_config_path, get_project_root
            # 如果传入的是相对路径，则基于项目根目录解析
            if not os.path.isabs(config_dir):
                self._config_dir = os.path.join(get_project_root(), config_dir)
            else:
                self._config_dir = os.path.normpath(config_dir)
        except ImportError:
            # 路径工具不可用时，使用备用方案确保绝对路径
            # 根据运行环境确定基础目录
            import sys
            if getattr(sys, 'frozen', False):
                # 打包环境：使用 exe 所在目录
                base_dir = os.path.dirname(sys.executable)
            else:
                # 开发环境：config_manager.py 在 src/docwen/config/ 下
                # 项目根目录 = config_manager.py 所在目录的上三级
                base_dir = os.path.dirname(os.path.dirname(os.path.dirname(
                    os.path.dirname(os.path.abspath(__file__)))))
            
            if not os.path.isabs(config_dir):
                self._config_dir = os.path.join(base_dir, config_dir)
            else:
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
    
    # ==========================================================================
    #                           配置文件路径获取
    # ==========================================================================
    
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
    
    # ==========================================================================
    #                           配置重载
    # ==========================================================================
    
    def reload_configs(self):
        """重新加载所有配置文件"""
        safe_log.info("重新加载所有配置文件...")
        self._initialized = False
        self._configs = {}
        self._initialize(self._config_dir)
    
    # ==========================================================================
    #                           配置修改方法
    # ==========================================================================
    
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
