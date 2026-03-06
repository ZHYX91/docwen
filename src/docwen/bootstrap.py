"""
应用统一启动器模块

本模块负责处理所有应用入口点（如GUI、CLI）共享的通用初始化任务。
主要功能包括：
- 运行核心安全检查（如反调试、反虚拟机）
- 初始化网络隔离功能
- 提供一个统一的初始化入口点

使用本模块可以确保所有启动流程的一致性，并避免在多个启动脚本中重复代码。
"""

import logging
import os
import sys
from pathlib import Path

# 获取根日志记录器（已由入口点预初始化）
logger = logging.getLogger()


def initialize_app(
    *, strict_security: bool | None = None, return_status: bool = False
) -> tuple[bool, str | None] | None:
    """
    执行所有应用共享的核心初始化任务。
    这个函数应该在任何特定于界面的代码运行之前被调用。
    """
    if strict_security is None:
        env_value = os.environ.get("DOCWEN_STRICT_SECURITY", "1").strip().lower()
        strict_security = env_value not in {"0", "false", "no", "off"}

    # 确保docwen包在Python路径中
    # 这一步对于从IDE直接运行脚本（如gui_run.py）至关重要
    package_dir = str(Path(__file__).resolve().parent.parent)
    if package_dir not in sys.path:
        sys.path.insert(0, package_dir)
        logger.debug(f"已将包目录添加到Python路径: {package_dir}")

    logger.info("=" * 50)
    logger.info("应用初始化开始...")

    # 1. 运行核心安全防护
    security_degraded_message: str | None = None
    try:
        from docwen.security.protection_utils import run_all_protections

        run_all_protections()
        logger.info("核心安全防护检查完成。")
    except ImportError as e:
        logger.warning(f"无法导入安全模块: {e}。程序将在无防护模式下运行。")
    except Exception as e:
        # 任何安全检查失败都应立即终止程序
        logger.critical(f"核心安全检查失败: {e}", exc_info=True)
        if strict_security:
            logger.critical("严格安全模式已启用，程序将终止。")
            from docwen.errors import SecurityCheckFailedError

            raise SecurityCheckFailedError(details=str(e), cause=e) from e
        security_degraded_message = f"核心安全检查失败，已降级运行: {e}"

    # 2. 初始化网络隔离
    try:
        from docwen.security.network_isolation import initialize_network_isolation

        if initialize_network_isolation():
            logger.info("网络隔离初始化成功 - 所有网络连接已被阻断。")
        else:
            logger.warning("网络隔离初始化部分失败，但程序将继续运行。")
    except ImportError as e:
        logger.warning(f"无法导入网络隔离模块: {e}。网络功能可能可用。")
    except Exception as e:
        logger.error(f"网络隔离初始化时发生未知错误: {e}")

    logger.info("应用初始化完成。")
    logger.info("=" * 50)
    status = (security_degraded_message is None, security_degraded_message)
    if return_status:
        return status
    return None
