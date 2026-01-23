"""
公文转换器CLI启动模块

统一的CLI入口，负责：
1. 日志系统初始化
2. 应用环境初始化
3. 有效期检查（CLI模式：仅终止程序）
4. 启动CLI主程序
"""

import logging
import sys

# 日志预初始化（必须最先执行）
from docwen.utils.logging_utils import pre_init_logging, init_logging_system
pre_init_logging()

from docwen.bootstrap import initialize_app

if __name__ == "__main__":
    logger = logging.getLogger()
    
    try:
        # 1. 应用初始化（路径、安全、网络隔离）
        logger.info("初始化应用环境...")
        initialize_app()
        
        # 2. 有效期检查（CLI模式：静默检查，过期则退出）
        logger.info("检查软件有效期...")
        try:
            from docwen.security.expiration_check import check_expiration
            check_expiration()
        except Exception as e:
            logger.critical(f"软件已过期或有效期检查失败: {e}")
            print("\n错误: 软件已过期，请联系管理员更新")
            sys.exit(1)
        
        # 3. 正式日志系统初始化
        logger.info("初始化日志系统...")
        init_logging_system()
        
        # 4. 启动CLI主程序
        logger.info("启动CLI主程序...")
        from docwen.cli.main import main
        sys.exit(main())  # 返回退出码
        
    except KeyboardInterrupt:
        logger.info("用户中断程序")
        print("\n程序已中断")
        sys.exit(130)  # 标准退出码：128+SIGINT(2)
        
    except Exception as e:
        logger.critical(f"CLI启动失败: {e}", exc_info=True)
        print(f"\n严重错误: {e}")
        print("请查看日志文件获取详细信息")
        sys.exit(1)
