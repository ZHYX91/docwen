"""
配置管理器

提供配置的加载、访问和修改功能，支持 GUI 配置和 TOML 文件持久化。

架构：
    - 核心类负责：配置加载、配置合并、配置修改
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

from pathlib import Path
from typing import Any

from .safe_logger import safe_log
from .schemas import (
    CONFIG_FILES,
    DEFAULT_CONFIG,
    ConversionConfigMixin,
    GUIConfigMixin,
    LinkConfigMixin,
    LoggerConfigMixin,
    OptimizationConfigMixin,
    OutputConfigMixin,
    ProofreadConfigMixin,
    SoftwareConfigMixin,
    StyleConfigMixin,
)
from .toml_operations import read_toml_file, update_toml_value


class ConfigManager(
    LoggerConfigMixin,
    GUIConfigMixin,
    ProofreadConfigMixin,
    OutputConfigMixin,
    SoftwareConfigMixin,
    LinkConfigMixin,
    StyleConfigMixin,
    ConversionConfigMixin,
    OptimizationConfigMixin,
):
    """
    配置管理器

    通过 Mixin 模式组合各功能域的配置获取方法。
    核心类负责配置的加载、合并和修改。
    """

    _configs: dict[str, dict[str, Any]]
    _config_dir = ""

    def __init__(self, config_dir: str | None = None) -> None:
        self._config_dir = self._resolve_config_dir(config_dir or "configs")
        self._configs = {}

        if not Path(self._config_dir).is_dir():
            safe_log.info("配置目录不存在: %s，使用默认配置", self._config_dir)
        self._load_all_configs()
        safe_log.info("配置管理器初始化完成 | 目录: %s", self._config_dir)

    @staticmethod
    def _resolve_config_dir(config_dir: str) -> str:
        try:
            from docwen.utils.path_utils import get_project_root

            if not Path(config_dir).is_absolute():
                return str(Path(get_project_root()) / config_dir)
            return str(Path(config_dir))
        except ImportError:
            import sys

            if getattr(sys, "frozen", False):
                base_dir = str(Path(sys.executable).resolve().parent)
            else:
                base_dir = str(Path(__file__).resolve().parents[3])
            if not Path(config_dir).is_absolute():
                return str(Path(base_dir) / config_dir)
            return str(Path(config_dir))

    def _load_all_configs(self):
        """加载所有配置文件，自动合并默认值"""
        loaded = 0
        for name, filename in CONFIG_FILES.items():
            self._reload_config_block(name)
            loaded += 1
            safe_log.debug("加载配置块: %s | 文件: %s", name, filename)
        safe_log.info("配置块加载完成: %d 个 | 目录: %s", loaded, self._config_dir)

    def _reload_config_block(self, config_name: str) -> None:
        filename = CONFIG_FILES[config_name]
        user_config = self._load_single_config(filename)
        default_config = DEFAULT_CONFIG.get(config_name, {})
        self._configs[config_name] = deep_merge_dicts(default_config, user_config)

    def _load_single_config(self, filename: str) -> dict[str, Any]:
        """安全加载单个配置文件"""
        filepath = str(Path(self._config_dir) / filename)

        if not Path(filepath).exists():
            safe_log.debug("配置文件不存在: %s", filepath)
            return {}

        user_config = read_toml_file(filepath)
        if not user_config:
            safe_log.warning("配置文件为空或格式错误: %s", filepath)
            return {}

        return user_config

    def reload_configs(self) -> None:
        safe_log.info("重新加载所有配置文件...")
        self._configs = {}
        self._load_all_configs()

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
        return str(Path(self._config_dir) / filename)

    # ==========================================================================
    #                           配置重载
    # ==========================================================================

    def reload_configs_from_dir(self, config_dir: str) -> None:
        self._config_dir = self._resolve_config_dir(config_dir)
        self.reload_configs()

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
        filepath = str(Path(self._config_dir) / filename)

        success = update_toml_value(filepath, section, key, value)
        if success:
            # 重新加载该配置文件以更新内存中的配置
            self._reload_config_block(config_name)
            safe_log.info("配置更新成功: %s -> %s.%s = %s", config_name, section, key, value)

        return success

    def update_config_section(self, config_name: str, section: str, data: dict[str, Any]) -> bool:
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
        filepath = str(Path(self._config_dir) / filename)

        try:
            # 读取现有配置（保留注释）
            from .toml_operations import read_toml_document, write_toml_document

            doc = read_toml_document(filepath)
            if doc is None:
                # 如果文件不存在或读取失败，创建新文档
                from tomlkit import document

                doc = document()

            # 分割多级节名称
            section_parts = section.split(".")

            # 导航到目标节
            current: Any = doc
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
                self._reload_config_block(config_name)
                safe_log.info("配置节更新成功: %s -> %s", config_name, section)

            return success

        except Exception as e:
            safe_log.error("更新配置节失败: %s -> %s | 错误: %s", config_name, section, str(e))
            return False


def deep_merge_dicts(default: dict, user: dict) -> dict:
    result = default.copy()
    for key, user_value in user.items():
        if key not in result:
            result[key] = user_value
        elif isinstance(result[key], dict) and isinstance(user_value, dict):
            result[key] = deep_merge_dicts(result[key], user_value)
        else:
            result[key] = user_value
    return result


config_manager = ConfigManager()
