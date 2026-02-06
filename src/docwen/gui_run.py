import ctypes
import logging
import os
import platform
import sys
from typing import Optional, Tuple

from docwen.bootstrap import initialize_app
from docwen.utils.gui_utils import show_error_dialog, show_info_dialog
from docwen.utils.logging_utils import init_logging_system, pre_init_logging

pre_init_logging()


def check_dependencies() -> Tuple[bool, Optional[str]]:
    logger = logging.getLogger()

    try:
        import tkinter  # noqa: F401
    except ImportError:
        return False, "Tkinter不可用，GUI无法启动。"

    try:
        import tkinterdnd2  # noqa: F401
    except ImportError:
        return False, "tkinterdnd2未安装，请运行: pip install tkinterdnd2。"

    try:
        import ttkbootstrap  # noqa: F401
    except ImportError:
        return False, "ttkbootstrap未安装，请运行: pip install ttkbootstrap。"

    logger.info("所有GUI依赖检查通过。")
    return True, None


def load_gui_config() -> bool:
    logger = logging.getLogger()

    try:
        from docwen.config.config_manager import config_manager

        gui_config = config_manager.get_gui_config_block()
        if not gui_config:
            logger.warning("GUI配置为空，使用默认设置")
            return False

        width = config_manager.get_center_panel_width()
        height = config_manager.get_window_height()
        theme = config_manager.get_default_theme()
        transparency = config_manager.get_transparency_value()

        logger.info(f"GUI配置加载成功: 尺寸={width}x{height}, 主题={theme}, 透明度={transparency}")
        return True
    except Exception as e:
        logger.error(f"加载GUI配置失败: {str(e)}")
        return False


def main() -> int:
    if platform.system() == "Windows":
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except (AttributeError, OSError):
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except (AttributeError, OSError):
                pass

    logger = logging.getLogger()

    from docwen.ipc.file_ipc import FileIPC
    from docwen.ipc.single_instance import SingleInstance

    instance_lock = SingleInstance("docwen")

    if not instance_lock.acquire():
        logger.info("检测到程序已在运行，将命令发送到已运行的实例")
        ipc_dir = instance_lock.get_ipc_dir()

        if len(sys.argv) > 1:
            file_path = sys.argv[1]
            command = {"action": "add_file", "file_path": file_path, "mode": "single"}
            success = FileIPC.send_command(ipc_dir, command)
            if success:
                logger.info(f"已将文件发送到运行中的程序: {file_path}")
                print(f"已将文件发送到运行中的程序: {file_path}")
            else:
                logger.error("发送文件命令失败")
                print("发送文件命令失败")
        else:
            command = {"action": "activate"}
            success = FileIPC.send_command(ipc_dir, command)
            if success:
                logger.info("已发送激活窗口命令")
                print("程序已在运行，已激活窗口")
            else:
                logger.error("发送激活命令失败")
                print("发送激活命令失败")

        return 0

    initialize_app()

    try:
        from docwen.security.expiration_check import ExpirationStatus, get_expiration_status

        expiration_status_info = get_expiration_status()

        deps_ok, error_msg = check_dependencies()
        if not deps_ok:
            show_error_dialog("依赖缺失", error_msg or "缺少必要的依赖包。")
            return 1

        load_gui_config()

        try:
            init_logging_system()
        except Exception as e:
            logger.critical(f"正式日志系统初始化失败: {e}", exc_info=True)

        try:
            import ttkbootstrap as tb
            from tkinterdnd2 import TkinterDnD

            root = TkinterDnD.Tk()
            style = tb.Style()
            root.style = style
            root.title("公文转换器")

            try:
                from docwen.utils.icon_utils import get_icon_path

                icon_path = get_icon_path()
                if icon_path and os.path.exists(icon_path):
                    root.iconbitmap(icon_path)
            except Exception:
                pass

            from docwen.config.config_manager import config_manager
            from docwen.gui.core.window import MainWindow

            initial_file = sys.argv[1] if len(sys.argv) > 1 else None
            app = MainWindow(root, config_manager=config_manager, initial_file_path=initial_file)

            def show_expiration_reminder() -> None:
                if expiration_status_info.status == ExpirationStatus.NEARING_EXPIRATION:
                    show_info_dialog(
                        "试用期提醒",
                        f"此软件的测试版本有效期不到3个月（剩余 {expiration_status_info.days_left} 天）。\n到期后，部分操作可能受限，请及时联系开发者。",
                    )
                elif expiration_status_info.status == ExpirationStatus.EXPIRED:
                    show_info_dialog(
                        "试用期已到期",
                        "此软件的测试版本已过期。\n请联系开发者获取正式版本。",
                        alert=True,
                    )

            root.after(200, show_expiration_reminder)

            ipc_dir = instance_lock.get_ipc_dir()
            file_ipc = FileIPC(ipc_dir, app.handle_ipc_command)
            file_ipc.start()

            try:
                root.mainloop()
            finally:
                file_ipc.stop()
                instance_lock.release()

            return 0
        except ImportError as e:
            show_error_dialog("导入错误", f"无法导入GUI模块: {str(e)}")
            return 1
        except Exception as e:
            show_error_dialog("启动错误", f"GUI启动失败: {str(e)}")
            return 1
    except Exception as e:
        logger.critical(f"GUI主程序崩溃: {str(e)}", exc_info=True)
        show_error_dialog("致命错误", f"程序崩溃: {str(e)}")
        return 1


if __name__ == "__main__":
    sys.exit(main())

