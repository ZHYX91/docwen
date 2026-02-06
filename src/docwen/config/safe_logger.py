"""
安全日志记录器

用于配置系统初始化前的日志记录，避免与日志系统的循环依赖问题。

详细说明：
    配置管理器在加载日志配置前需要记录日志，但此时日志系统尚未初始化。
    SafeLogger 提供一个安全的日志接口，在日志系统就绪前使用标准输出，
    就绪后自动切换到标准日志系统。

特性：
    - 不依赖标准日志系统，避免循环依赖
    - 支持在日志系统初始化前记录日志
    - 自动降级到标准输出/错误输出
    - 可禁用日志记录（用于测试）

依赖：
    - logging: Python 标准日志模块
    - sys: 标准输入输出

使用方式：
    # 使用全局实例
    from docwen.config.safe_logger import safe_log
    
    safe_log.info("配置加载完成: %s", config_name)
    safe_log.error("配置解析失败: %s", error_message)
    
    # 使用模块级函数
    from docwen.config.safe_logger import info, error
    
    info("这是一条信息")
    error("这是一条错误")
"""

import sys
import logging
from typing import Optional


class SafeLogger:
    """
    安全日志记录器，用于配置系统初始化前的日志记录
    
    特性：
        1. 不依赖标准日志系统，避免循环依赖
        2. 支持在日志系统初始化前记录日志
        3. 自动降级到标准输出/错误输出
        4. 可禁用日志记录（用于测试）
    
    属性：
        _enabled: 是否启用日志记录
        _logger: 底层 logging.Logger 实例（延迟初始化）
        _name: 日志记录器名称
    """
    
    def __init__(self) -> None:
        """初始化安全日志记录器"""
        self._enabled: bool = True
        self._logger: Optional[logging.Logger] = None
        self._name: str = "config_manager"
    
    def _setup_logger(self) -> None:
        """
        设置日志记录器（延迟初始化）
        
        在首次调用日志方法时初始化底层 Logger，
        如果初始化失败则保持 _logger 为 None，后续使用标准输出。
        """
        if self._logger is None:
            try:
                self._logger = logging.getLogger(self._name)
                # 设置默认处理器，避免"No handlers"警告
                if not self._logger.handlers:
                    handler = logging.StreamHandler(sys.stderr)
                    formatter = logging.Formatter(
                        "%(asctime)s [%(levelname)s] %(name)s:%(lineno)d - %(message)s"
                    )
                    handler.setFormatter(formatter)
                    self._logger.addHandler(handler)
                    self._logger.setLevel(logging.INFO)
            except Exception:
                # 如果设置失败，保持 self._logger 为 None
                self._logger = None
    
    def log(self, level: str, message: str, *args) -> None:
        """
        安全记录日志
        
        参数：
            level: 日志级别，可选值: debug/info/warning/error/critical
            message: 日志消息，可以包含 % 格式化占位符
            *args: 格式化参数，用于替换 message 中的占位符
        
        返回：
            None
        
        示例：
            log("info", "处理文件: %s", filename)
            log("error", "错误代码: %d, 消息: %s", code, msg)
        """
        if not self._enabled:
            return
            
        # 格式化消息
        formatted_message = message % args if args else message
        formatted_output = f"[{self._name}] {level.upper()}: {formatted_message}"
        
        try:
            self._setup_logger()
            
            # 尝试使用标准日志系统
            if self._logger is not None:
                log_method = getattr(self._logger, level, None)
                if log_method and callable(log_method):
                    log_method(formatted_message)
                    return
                
            # 回退到简单打印
            if level in ("error", "critical", "warning"):
                print(formatted_output, file=sys.stderr)
            else:
                print(formatted_output)
                
        except Exception:
            # 最终回退到简单打印
            if level in ("error", "critical", "warning"):
                print(formatted_output, file=sys.stderr)
            else:
                print(formatted_output)
    
    def debug(self, message: str, *args) -> None:
        """
        记录调试信息
        
        参数：
            message: 日志消息，可以包含 % 格式化占位符
            *args: 格式化参数
        
        返回：
            None
        """
        self.log("debug", message, *args)
    
    def info(self, message: str, *args) -> None:
        """
        记录一般信息
        
        参数：
            message: 日志消息，可以包含 % 格式化占位符
            *args: 格式化参数
        
        返回：
            None
        """
        self.log("info", message, *args)
    
    def warning(self, message: str, *args) -> None:
        """
        记录警告信息
        
        参数：
            message: 日志消息，可以包含 % 格式化占位符
            *args: 格式化参数
        
        返回：
            None
        """
        self.log("warning", message, *args)
    
    def error(self, message: str, *args) -> None:
        """
        记录错误信息
        
        参数：
            message: 日志消息，可以包含 % 格式化占位符
            *args: 格式化参数
        
        返回：
            None
        """
        self.log("error", message, *args)
    
    def critical(self, message: str, *args) -> None:
        """
        记录严重错误信息
        
        参数：
            message: 日志消息，可以包含 % 格式化占位符
            *args: 格式化参数
        
        返回：
            None
        """
        self.log("critical", message, *args)
    
    def disable(self) -> None:
        """
        禁用日志记录
        
        用于测试场景，禁用后所有日志调用将被忽略。
        
        返回：
            None
        """
        self._enabled = False
    
    def enable(self) -> None:
        """
        启用日志记录
        
        返回：
            None
        """
        self._enabled = True
    
    def is_enabled(self) -> bool:
        """
        检查日志记录是否启用
        
        返回：
            bool: True 表示已启用，False 表示已禁用
        """
        return self._enabled
    
    def set_name(self, name: str) -> None:
        """
        设置日志记录器名称
        
        参数：
            name: 新的日志记录器名称
        
        返回：
            None
        
        注意：
            设置后会重置内部 logger 实例，下次调用时重新初始化
        """
        self._name = name
        # 重置 logger 以便下次使用时重新初始化
        self._logger = None


# ==============================================================================
#                              全局实例
# ==============================================================================

# 创建全局安全日志记录器实例
safe_log = SafeLogger()


# ==============================================================================
#                              模块级函数
# ==============================================================================

def debug(message: str, *args) -> None:
    """
    记录调试信息（模块级函数）
    
    参数：
        message: 日志消息
        *args: 格式化参数
    
    返回：
        None
    """
    safe_log.debug(message, *args)


def info(message: str, *args) -> None:
    """
    记录一般信息（模块级函数）
    
    参数：
        message: 日志消息
        *args: 格式化参数
    
    返回：
        None
    """
    safe_log.info(message, *args)


def warning(message: str, *args) -> None:
    """
    记录警告信息（模块级函数）
    
    参数：
        message: 日志消息
        *args: 格式化参数
    
    返回：
        None
    """
    safe_log.warning(message, *args)


def error(message: str, *args) -> None:
    """
    记录错误信息（模块级函数）
    
    参数：
        message: 日志消息
        *args: 格式化参数
    
    返回：
        None
    """
    safe_log.error(message, *args)


def critical(message: str, *args) -> None:
    """
    记录严重错误信息（模块级函数）
    
    参数：
        message: 日志消息
        *args: 格式化参数
    
    返回：
        None
    """
    safe_log.critical(message, *args)


def disable() -> None:
    """
    禁用日志记录（模块级函数）
    
    返回：
        None
    """
    safe_log.disable()


def enable() -> None:
    """
    启用日志记录（模块级函数）
    
    返回：
        None
    """
    safe_log.enable()
