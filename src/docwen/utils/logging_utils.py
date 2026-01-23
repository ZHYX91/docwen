"""
统一日志系统工具模块
实现功能：
1. 根据配置初始化日志系统
2. 按天轮转日志文件（docwen_20250910.log）
3. 支持文件和控制台双输出
4. 自动清理过期日志
5. 安全错误处理机制
6. 禁用极简日志系统
"""

import os
import sys
import logging
import re
import time
import datetime
from logging.handlers import TimedRotatingFileHandler

from docwen.config.config_manager import config_manager

# 全局标志，防止重复初始化
_logging_initialized = False

def pre_init_logging() -> logging.Logger:
    """
    预初始化日志系统（在配置加载前调用）
    - 配置根日志记录器
    - 添加一个极简的控制台处理器，用于捕获所有模块的早期日志
    """
    root_logger = logging.getLogger()
    
    # 如果已经有处理器，可能在测试或重复调用中，先清除
    if root_logger.handlers:
        for handler in root_logger.handlers[:]:
            root_logger.removeHandler(handler)
            
    # 设置一个基础级别
    root_logger.setLevel(logging.DEBUG)

    # 创建一个极简的控制台处理器
    console_handler = logging.StreamHandler(sys.stderr)
    formatter = logging.Formatter("[%(levelname).1s] %(name)s: %(message)s")
    console_handler.setFormatter(formatter)
    
    root_logger.addHandler(console_handler)
    
    root_logger.info("日志系统预初始化完成。")
    return root_logger

def generate_log_path() -> str:
    """
    生成日志文件路径（安全方法）
    开发环境：项目根目录/logs/
    生产环境：软件根目录/logs/
    格式：<前缀>_<日期>.log（如docwen_20250910.log）
    
    参数:
        config: 日志配置字典
        
    返回:
        str: 完整的日志文件路径
    """
    # 检测开发环境（通过检查是否在 src 目录下运行）
    script_path = os.path.abspath(sys.argv[0])
    if "src" in script_path and os.path.basename(os.path.dirname(script_path)) == "src":
        # 开发环境：使用项目根目录
        base_dir = os.path.dirname(os.path.dirname(script_path))
    else:
        # 生产环境：使用软件根目录
        base_dir = os.path.dirname(script_path)
    
    # 获取日志目录
    log_dir = os.path.join(base_dir, "logs")
    
    # 确保日志目录存在
    try:
        os.makedirs(log_dir, exist_ok=True)
        logging.getLogger().debug(f"创建/确认日志目录: {log_dir}")
    except Exception as e:
        # 回退到临时目录
        import tempfile
        log_dir = tempfile.gettempdir()
        logging.getLogger().warning(f"创建日志目录失败，使用临时目录: {log_dir}")
    
    # 生成文件名
    config = config_manager.get_logging_config()
    date_str = datetime.datetime.now().strftime("%Y%m%d")
    prefix = config.get("file_prefix", "docwen")
    
    # 清理文件名中的非法字符
    safe_prefix = re.sub(r'[\\/*?:"<>|]', "", prefix)
    filename = f"{safe_prefix}_{date_str}.log"
    
    full_path = os.path.join(log_dir, filename)
    logging.getLogger().debug(f"生成日志路径: {full_path}")
    return full_path

def configure_file_handler(logger: logging.Logger) -> bool:
    """
    配置文件日志处理器
    
    参数:
        logger: 日志记录器对象
        
    返回:
        bool: 是否成功配置
    """
    try:
        config = config_manager.get_logging_config()
        # 创建按天轮转的日志处理器
        file_handler = TimedRotatingFileHandler(
            filename=generate_log_path(),
            when="midnight",  # 每天午夜轮转
            interval=1,       # 每天一次
            backupCount=config.get("retention_days", 7),
            encoding='utf-8'
        )
        
        # 设置日志格式（硬编码）
        formatter = logging.Formatter(
            "%(asctime)s [%(levelname)s] %(name)s:%(lineno)d - %(message)s"
        )
        file_handler.setFormatter(formatter)
        
        # 设置日志级别
        file_level = config.get("level", "info").upper()
        level_mapping = {
            "DEBUG": logging.DEBUG,
            "INFO": logging.INFO,
            "WARNING": logging.WARNING,
            "ERROR": logging.ERROR,
            "CRITICAL": logging.CRITICAL
        }
        file_handler.setLevel(level_mapping.get(file_level, logging.INFO))
        
        # 添加到日志记录器
        logger.addHandler(file_handler)
        logging.getLogger().info(f"文件日志处理器配置成功，级别: {file_level}")
        return True
        
    except PermissionError:
        # 权限问题处理
        logging.getLogger().critical("没有日志文件写入权限，文件日志功能禁用")
        return False
    except Exception as e:
        # 其他错误处理
        logging.getLogger().critical(f"文件日志配置失败: {str(e)}")
        return False

def configure_console_handler(logger: logging.Logger) -> bool:
    """
    配置控制台日志处理器
    
    参数:
        logger: 日志记录器对象
        
    返回:
        bool: 是否成功配置
    """
    try:
        config = config_manager.get_logging_config()
        # 创建控制台处理器
        console_handler = logging.StreamHandler(sys.stdout)
        
        # 设置日志格式（硬编码）
        formatter = logging.Formatter(
            "%(asctime)s [%(levelname)s] %(name)s:%(lineno)d - %(message)s"
        )
        console_handler.setFormatter(formatter)
        
        # 设置控制台日志级别
        console_level = config.get("console_level", "info").upper()
        level_mapping = {
            "DEBUG": logging.DEBUG,
            "INFO": logging.INFO,
            "WARNING": logging.WARNING,
            "ERROR": logging.ERROR,
            "CRITICAL": logging.CRITICAL
        }
        console_handler.setLevel(level_mapping.get(console_level, logging.INFO))
        
        # 添加到日志记录器
        logger.addHandler(console_handler)
        logging.getLogger().info(f"控制台日志处理器配置成功，级别: {console_level}")
        return True
        
    except Exception as e:
        # 错误处理
        logging.getLogger().error(f"控制台日志配置失败: {str(e)}")
        return False

def init_logging_system() -> logging.Logger:
    """
    初始化完整日志系统（主入口函数）
    此函数现在负责正式的、基于配置的初始化。
        
    返回:
        logging.Logger: 配置好的根日志记录器
    """
    global _logging_initialized
    if _logging_initialized:
        logging.getLogger().warning("日志系统已初始化，跳过重复的正式初始化。")
        return logging.getLogger()
    
    config = config_manager.get_logging_config()
    # 获取根日志记录器
    root_logger = logging.getLogger()
    root_logger.handlers = []  # 清除所有现有处理器（包括预初始化的）
    
    # 设置日志级别
    log_level = config.get("level", "info").upper()
    level_mapping = {
        "DEBUG": logging.DEBUG,
        "INFO": logging.INFO,
        "WARNING": logging.WARNING,
        "ERROR": logging.ERROR,
        "CRITICAL": logging.CRITICAL
    }
    root_logger.setLevel(level_mapping.get(log_level, logging.INFO))
    
    # 配置文件日志（如果启用）
    if config.get("enable", True):
        success = configure_file_handler(root_logger)
        if not success:
            # 文件日志失败时禁用文件日志
            config["enable"] = False
    
    # 配置控制台日志（如果启用）
    if config.get("console_enable", True):
        configure_console_handler(root_logger)
    
    # 设置初始化标志
    _logging_initialized = True
    
    # 记录初始化完成消息
    log_handlers = ", ".join([type(h).__name__ for h in root_logger.handlers])
    root_logger.critical("=" * 80)
    root_logger.critical("日志系统初始化完成")
    root_logger.critical(f"日志级别: {log_level}")
    root_logger.critical(f"日志处理器: {log_handlers}")
    root_logger.critical(f"日志保留天数: {config.get('retention_days', 7)}")
    root_logger.critical("=" * 80)
    
    return root_logger

def clean_old_logs() -> None:
    """
    清理过期日志文件（按计划自动执行）
    """
    try:
        config = config_manager.get_logging_config()
        # 获取日志目录
        log_dir = os.path.dirname(generate_log_path())
        
        # 计算过期时间点
        retention_days = config.get("retention_days", 7)
        cutoff_time = time.time() - (retention_days * 86400)
        
        # 遍历日志目录
        for filename in os.listdir(log_dir):
            file_path = os.path.join(log_dir, filename)
            
            # 检查文件是否过期
            if os.path.isfile(file_path) and os.stat(file_path).st_mtime < cutoff_time:
                try:
                    os.remove(file_path)
                    logging.info(f"清理过期日志: {filename}")
                except Exception as e:
                    logging.warning(f"清理日志失败 {filename}: {str(e)}")
                    
    except Exception as e:
        logging.error(f"日志清理失败: {str(e)}")

# 模块测试代码
if __name__ == "__main__":
    
    print("--- 预初始化测试 ---")
    pre_init_logging()
    logging.getLogger("test.module").info("这条消息应该通过预初始化日志显示")

    print("\n--- 正式初始化测试 ---")
    # 初始化日志系统
    logger = init_logging_system()
    
    # 测试日志记录
    logger.debug("调试信息")
    logger.info("普通信息")
    logger.warning("警告信息")
    logger.error("错误信息")
    logger.critical("严重错误")
    
    # 测试日志清理
    logger.info("测试日志清理...")
    clean_old_logs()
    
    logger.info("日志系统测试完成")
