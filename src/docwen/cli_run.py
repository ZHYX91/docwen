import contextlib
import sys


def _early_reconfigure_stdio() -> None:
    try:
        if sys.platform == "win32":
            try:
                import ctypes

                kernel32 = ctypes.windll.kernel32
                kernel32.SetConsoleOutputCP(65001)
                kernel32.SetConsoleCP(65001)
            except Exception:
                pass

        try:
            import io

            stdout_reconfigure = getattr(sys.stdout, "reconfigure", None)
            stderr_reconfigure = getattr(sys.stderr, "reconfigure", None)
            if callable(stdout_reconfigure):
                with contextlib.suppress(Exception):
                    stdout_reconfigure(encoding="utf-8", errors="replace")
            if callable(stderr_reconfigure):
                with contextlib.suppress(Exception):
                    stderr_reconfigure(encoding="utf-8", errors="replace")

            if isinstance(sys.stdout, io.TextIOWrapper) and getattr(sys.stdout, "encoding", None) != "utf-8":
                sys.stdout = io.TextIOWrapper(
                    sys.stdout.buffer, encoding="utf-8", errors="replace", line_buffering=True
                )
            if isinstance(sys.stderr, io.TextIOWrapper) and getattr(sys.stderr, "encoding", None) != "utf-8":
                sys.stderr = io.TextIOWrapper(
                    sys.stderr.buffer, encoding="utf-8", errors="replace", line_buffering=True
                )
        except Exception:
            pass
    except Exception:
        pass


_early_reconfigure_stdio()

import logging  # noqa: E402

from docwen.bootstrap import initialize_app  # noqa: E402
from docwen.errors import DocWenError, ExitCode, SecurityCheckFailedError  # noqa: E402
from docwen.services.error_codes import (  # noqa: E402
    ERROR_CODE_DEPENDENCY_MISSING,
    ERROR_CODE_INVALID_INPUT,
    ERROR_CODE_STRATEGY_NOT_FOUND,
)
from docwen.utils.logging_utils import init_logging_system, pre_init_logging  # noqa: E402

pre_init_logging()


def main() -> int:
    logger = logging.getLogger()

    try:
        logger.info("初始化应用环境...")
        init_status = initialize_app(return_status=True)
        init_ok, init_msg = init_status if init_status else (True, None)
        if not init_ok and init_msg:
            logger.warning(init_msg)
            print(f"\n警告: {init_msg}", file=sys.stderr)

        logger.info("初始化日志系统...")
        init_logging_system()

        logger.info("启动CLI主程序...")
        from docwen.cli.main import main as cli_main

        return int(cli_main())
    except KeyboardInterrupt:
        logger.info("用户中断程序")
        print("\n程序已中断", file=sys.stderr)
        return 130
    except SecurityCheckFailedError as e:
        logger.critical(f"启动安全检查失败: {e}", exc_info=True)
        print(f"\n严重错误: {e}", file=sys.stderr)
        return int(ExitCode.UNKNOWN_ERROR)
    except DocWenError as e:
        logger.error(f"CLI运行失败: {e}", exc_info=True)
        print(f"\n错误: {e}", file=sys.stderr)
        if e.code == ERROR_CODE_INVALID_INPUT:
            return int(ExitCode.INVALID_INPUT)
        if e.code == ERROR_CODE_STRATEGY_NOT_FOUND:
            return int(ExitCode.STRATEGY_NOT_FOUND)
        if e.code == ERROR_CODE_DEPENDENCY_MISSING:
            return int(ExitCode.DEPENDENCY_MISSING)
        return int(ExitCode.UNKNOWN_ERROR)
    except Exception as e:
        logger.critical(f"CLI启动失败: {e}", exc_info=True)
        print(f"\n严重错误: {e}", file=sys.stderr)
        print("请查看日志文件获取详细信息", file=sys.stderr)
        return int(ExitCode.UNKNOWN_ERROR)


if __name__ == "__main__":
    sys.exit(main())
