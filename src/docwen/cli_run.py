import logging
import sys

from docwen.bootstrap import initialize_app
from docwen.utils.logging_utils import init_logging_system, pre_init_logging

pre_init_logging()


def main() -> int:
    logger = logging.getLogger()

    try:
        logger.info("初始化应用环境...")
        initialize_app()

        logger.info("检查软件有效期...")
        try:
            from docwen.security.expiration_check import check_expiration

            check_expiration()
        except Exception as e:
            logger.critical(f"软件已过期或有效期检查失败: {e}")
            print("\n错误: 软件已过期，请联系管理员更新")
            return 1

        logger.info("初始化日志系统...")
        init_logging_system()

        logger.info("启动CLI主程序...")
        from docwen.cli.main import main as cli_main

        return int(cli_main())
    except KeyboardInterrupt:
        logger.info("用户中断程序")
        print("\n程序已中断")
        return 130
    except Exception as e:
        logger.critical(f"CLI启动失败: {e}", exc_info=True)
        print(f"\n严重错误: {e}")
        print("请查看日志文件获取详细信息")
        return 1


if __name__ == "__main__":
    sys.exit(main())

