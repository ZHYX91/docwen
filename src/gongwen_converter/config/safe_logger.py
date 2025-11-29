"""
安全日志记录器
用于配置系统初始化前的日志记录，避免与日志系统的循环依赖
"""

import sys
import logging
from typing import Optional

class SafeLogger:
    """
    安全日志记录器，用于配置系统初始化前的日志记录
    
    特性:
    1. 不依赖标准日志系统，避免循环依赖
    2. 支持在日志系统初始化前记录日志
    3. 自动降级到标准输出/错误输出
    4. 可禁用日志记录（用于测试）
    """
    
    def __init__(self):
        self._enabled = True
        self._logger: Optional[logging.Logger] = None
        self._name = "config_manager"
    
    def _setup_logger(self):
        """设置日志记录器（延迟初始化）"""
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
                # 如果设置失败，保持self._logger为None
                self._logger = None
    
    def log(self, level: str, message: str, *args):
        """
        安全记录日志
        
        参数:
            level: 日志级别 (debug/info/warning/error/critical)
            message: 日志消息，可以包含格式化占位符
            *args: 格式化参数
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
    
    def debug(self, message: str, *args):
        """记录调试信息"""
        self.log("debug", message, *args)
    
    def info(self, message: str, *args):
        """记录一般信息"""
        self.log("info", message, *args)
    
    def warning(self, message: str, *args):
        """记录警告信息"""
        self.log("warning", message, *args)
    
    def error(self, message: str, *args):
        """记录错误信息"""
        self.log("error", message, *args)
    
    def critical(self, message: str, *args):
        """记录严重错误信息"""
        self.log("critical", message, *args)
    
    def disable(self):
        """禁用日志记录（用于测试）"""
        self._enabled = False
    
    def enable(self):
        """启用日志记录"""
        self._enabled = True
    
    def is_enabled(self) -> bool:
        """检查日志记录是否启用"""
        return self._enabled
    
    def set_name(self, name: str):
        """设置日志记录器名称"""
        self._name = name
        # 重置logger以便下次使用时重新初始化
        self._logger = None

# 创建全局安全日志记录器实例
safe_log = SafeLogger()

# 导出常用函数作为模块级函数
def debug(message: str, *args):
    """记录调试信息（模块级函数）"""
    safe_log.debug(message, *args)

def info(message: str, *args):
    """记录一般信息（模块级函数）"""
    safe_log.info(message, *args)

def warning(message: str, *args):
    """记录警告信息（模块级函数）"""
    safe_log.warning(message, *args)

def error(message: str, *args):
    """记录错误信息（模块级函数）"""
    safe_log.error(message, *args)

def critical(message: str, *args):
    """记录严重错误信息（模块级函数）"""
    safe_log.critical(message, *args)

def disable():
    """禁用日志记录（模块级函数）"""
    safe_log.disable()

def enable():
    """启用日志记录（模块级函数）"""
    safe_log.enable()

# 测试代码
if __name__ == "__main__":
    print("安全日志记录器测试开始...")
    
    # 测试各种日志级别
    debug("这是一条调试消息: %s", "调试信息")
    info("这是一条信息消息: %d", 42)
    warning("这是一条警告消息")
    error("这是一条错误消息: %s", "错误详情")
    critical("这是一条严重错误消息")
    
    # 测试禁用功能
    print("\n测试禁用功能...")
    disable()
    info("这条消息不应该显示")
    
    # 测试启用功能
    print("测试启用功能...")
    enable()
    info("这条消息应该显示")
    
    # 测试名称设置
    print("\n测试名称设置...")
    safe_log.set_name("test_logger")
    info("使用新名称的记录器")
    
    print("安全日志记录器测试完成!")
