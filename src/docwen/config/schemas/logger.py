"""
日志配置模块

对应配置文件：logger_config.toml

详细说明：
    包含日志系统的默认配置和访问方法。
    日志配置控制应用程序的日志行为，包括日志级别、文件输出、保留天数等。

包含：
    - DEFAULT_LOGGING_CONFIG: 默认日志配置
    - LoggerConfigMixin: 日志配置获取方法

依赖：
    - safe_logger: 安全日志记录（用于配置访问时的调试日志）

使用方式：
    # 通过 ConfigManager 访问（推荐）
    from docwen.config import config_manager

    log_level = config_manager.get_logging_config().get("level")

    # 直接导入默认配置
    from docwen.config.schemas.logger import DEFAULT_LOGGING_CONFIG
"""

from typing import Any

from ..safe_logger import safe_log

# ==============================================================================
#                              默认配置
# ==============================================================================

DEFAULT_LOGGING_CONFIG = {
    "logging": {
        "enable": True,
        "level": "debug",
        "file_prefix": "docwen",
        "retention_days": 30,
        "console_enable": True,
        "console_level": "info",
    }
}

# 配置文件名
CONFIG_FILE = "logger_config.toml"


# ==============================================================================
#                              Mixin 类
# ==============================================================================


class LoggerConfigMixin:
    """
    日志配置获取方法 Mixin

    提供日志相关配置的访问方法。

    注意：
        此类设计为 Mixin，需要与 ConfigManager 一起使用。
        假定宿主类具有 _configs 属性（配置字典）。

    配置结构：
        logger_config:
            logging:
                enable: bool - 是否启用日志
                level: str - 日志级别 (debug/info/warning/error)
                file_prefix: str - 日志文件前缀
                retention_days: int - 日志保留天数
                console_enable: bool - 是否启用控制台输出
                console_level: str - 控制台日志级别
    """

    # 类型提示：声明 _configs 属性（由 ConfigManager 提供）
    _configs: dict[str, dict[str, Any]]

    # --------------------------------------------------------------------------
    # 第一层：配置块
    # --------------------------------------------------------------------------

    def get_logger_config_block(self) -> dict[str, Any]:
        """
        获取整个日志配置块

        返回：
            Dict[str, Any]: 日志配置字典，包含 logging 子表

        示例：
            {
                "logging": {
                    "enable": True,
                    "level": "info",
                    ...
                }
            }
        """
        return self._configs.get("logger_config", {})

    # --------------------------------------------------------------------------
    # 第二层：子表
    # --------------------------------------------------------------------------

    def get_logging_config(self) -> dict[str, Any]:
        """
        获取日志配置子表

        返回：
            Dict[str, Any]: 日志配置子表，包含 enable、level 等字段

        示例：
            {
                "enable": True,
                "level": "info",
                "file_prefix": "docwen",
                "retention_days": 7,
                "console_enable": True,
                "console_level": "info"
            }
        """
        return self.get_logger_config_block().get("logging", {})

    # --------------------------------------------------------------------------
    # 第三层：具体配置值
    # --------------------------------------------------------------------------

    def is_logging_enabled(self) -> bool:
        """
        检查是否启用日志记录

        返回：
            bool: True 表示启用，False 表示禁用
        """
        logging_config = self.get_logging_config()
        enabled = logging_config.get("enable", True)
        safe_log.debug("日志启用状态: %s", enabled)
        return enabled

    def get_log_level(self) -> str:
        """
        获取日志级别

        返回：
            str: 日志级别，可选值: "debug", "info", "warning", "error"
        """
        logging_config = self.get_logging_config()
        level = logging_config.get("level", "info")
        # 确保返回有效的级别
        valid_levels = ["debug", "info", "warning", "error", "critical"]
        if level.lower() not in valid_levels:
            level = "info"
        safe_log.debug("获取日志级别: %s", level)
        return level.lower()

    def get_log_file_prefix(self) -> str:
        """
        获取日志文件前缀

        返回：
            str: 日志文件名前缀，如 "docwen" 会生成 "docwen_2024-01-01.log"
        """
        logging_config = self.get_logging_config()
        prefix = logging_config.get("file_prefix", "docwen")
        safe_log.debug("获取日志文件前缀: %s", prefix)
        return prefix

    def get_log_retention_days(self) -> int:
        """
        获取日志保留天数

        返回：
            int: 日志文件保留天数，超过此天数的旧日志将被清理
        """
        logging_config = self.get_logging_config()
        days = logging_config.get("retention_days", 7)
        # 确保在有效范围内
        if not isinstance(days, int) or days < 1:
            days = 7
        safe_log.debug("获取日志保留天数: %d", days)
        return days

    def is_console_logging_enabled(self) -> bool:
        """
        检查是否启用控制台日志输出

        返回：
            bool: True 表示启用控制台输出，False 表示禁用
        """
        logging_config = self.get_logging_config()
        enabled = logging_config.get("console_enable", True)
        safe_log.debug("控制台日志启用状态: %s", enabled)
        return enabled

    def get_console_log_level(self) -> str:
        """
        获取控制台日志级别

        返回：
            str: 控制台日志级别，可选值: "debug", "info", "warning", "error"
        """
        logging_config = self.get_logging_config()
        level = logging_config.get("console_level", "info")
        # 确保返回有效的级别
        valid_levels = ["debug", "info", "warning", "error", "critical"]
        if level.lower() not in valid_levels:
            level = "info"
        safe_log.debug("获取控制台日志级别: %s", level)
        return level.lower()
