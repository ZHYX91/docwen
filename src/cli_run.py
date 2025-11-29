"""
公文转换器命令行界面启动模块

本模块是公文转换器CLI应用程序的开发时运行入口，提供简化的启动流程。
主要功能包括：
- 初始化极简日志系统用于启动阶段错误记录
- 初始化网络隔离功能
- 调用CLI主程序入口

此模块主要用于开发调试阶段，生产环境建议使用打包后的可执行文件。
"""

import logging

# 统一日志系统: 在所有其他导入之前进行预初始化
from gongwen_converter.utils.logging_utils import pre_init_logging, init_logging_system
pre_init_logging()

from gongwen_converter.bootstrap import initialize_app

if __name__ == "__main__":
    logger = logging.getLogger()
    
    # 1. 运行统一的应用初始化
    # 这会处理路径、安全检查和网络隔离
    initialize_app()
    
    # 2. 运行有效期检查
    # 注意: CLI版本目前不处理弹窗，仅在过期时终止
    try:
        from gongwen_converter.security.expiration_check import check_expiration
        check_expiration()
    except Exception as e:
        logger.critical(f"有效期检查失败，程序无法启动: {e}")
        exit(1)

    # 3. 初始化正式日志系统 (在CLI主程序之前)
    try:
        init_logging_system()
        logger.info("正式日志系统初始化成功。")
    except Exception as e:
        logger.critical(f"正式日志系统初始化失败: {e}", exc_info=True)
        # 即使失败，程序仍可依赖预初始化日志继续运行

    # 4. 运行CLI主程序
    logger.info("启动CLI主程序...")
    from gongwen_converter.cli.main import main
    main()
