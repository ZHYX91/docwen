"""
高级保护工具模块
提供反调试功能
"""

import logging
import os
import sys

logger = logging.getLogger()


def check_debugger():
    """
    检测是否处于调试模式下。
    结合使用sys.gettrace()和Windows API IsDebuggerPresent()进行双重检查。
    """
    try:
        # 方法一: Python层面的追踪器检查
        if sys.gettrace() is not None:
            logger.critical("检测到调试器活动 (sys.gettrace)，程序将立即终止。")
            os._exit(1)

        # 方法二: Windows API层面的检查 (仅限Windows)
        if sys.platform == "win32":
            import ctypes

            is_debugger_present = ctypes.windll.kernel32.IsDebuggerPresent()
            if is_debugger_present:
                logger.critical("检测到调试器活动 (IsDebuggerPresent)，程序将立即终止。")
                os._exit(1)

    except Exception:
        # 在极端情况下，如果检查失败，也选择退出以策安全
        os._exit(1)


def run_all_protections():
    """
    运行所有高级保护检查。
    """
    check_debugger()
