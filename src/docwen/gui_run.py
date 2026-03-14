import contextlib
import ctypes
import faulthandler
import importlib.util
import logging
import os
import sys
import tempfile
import time
from pathlib import Path
from typing import TextIO

from docwen.bootstrap import initialize_app
from docwen.errors import SecurityCheckFailedError
from docwen.utils.gui_utils import show_error_dialog
from docwen.utils.logging_utils import init_logging_system, pre_init_logging

pre_init_logging()


def check_dependencies() -> tuple[bool, str | None]:
    logger = logging.getLogger()

    if importlib.util.find_spec("tkinter") is None:
        return False, "Tkinter不可用，GUI无法启动。"

    if importlib.util.find_spec("tkinterdnd2") is None:
        return False, "tkinterdnd2未安装，请运行: pip install tkinterdnd2。"

    if importlib.util.find_spec("ttkbootstrap") is None:
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
        logger.error(f"加载GUI配置失败: {e!s}")
        return False


def _setup_gui_hang_diagnostics(root) -> tuple[str, TextIO | None]:
    def _resolve_dump_dir() -> str:
        env_dir = os.environ.get("DOCWEN_GUI_HANG_DIR")
        if env_dir:
            return str(Path(os.path.expandvars(env_dir)).expanduser().resolve())
        return str(Path(tempfile.gettempdir()) / "docwen")

    def _cleanup_old_dumps(dir_path: str, keep: int = 10, max_age_days: int = 7) -> None:
        try:
            entries = []
            for path in Path(dir_path).iterdir():
                if not (path.name.startswith("gui_hang_") and path.name.endswith(".log")):
                    continue
                try:
                    st = path.stat()
                except OSError:
                    continue
                entries.append((st.st_mtime, path))
            if not entries:
                return
            entries.sort(reverse=True)
            cutoff_ts = time.time() - (max_age_days * 86400)
            for idx, (mtime, path) in enumerate(entries):
                if idx < keep and mtime >= cutoff_ts:
                    continue
                with contextlib.suppress(OSError):
                    path.unlink()
        except Exception:
            return

    dump_dir = _resolve_dump_dir()
    try:
        Path(dump_dir).mkdir(parents=True, exist_ok=True)
    except Exception:
        dump_dir = str(Path(tempfile.gettempdir()) / "docwen")
        Path(dump_dir).mkdir(parents=True, exist_ok=True)
    _cleanup_old_dumps(dump_dir)
    dump_path = str(Path(dump_dir) / f"gui_hang_{os.getpid()}_{int(time.time())}.log")
    fh: TextIO | None = None
    try:
        fh = Path(dump_path).open("a", encoding="utf-8")
        try:
            fh.write("=" * 80 + "\n")
            fh.write(f"start {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            fh.write(f"pid {os.getpid()}\n")
            fh.write(f"python {sys.version}\n")
            fh.write(f"argv {sys.argv}\n")
            fh.write("faulthandler dump_traceback_later: 60s repeat\n")
            fh.flush()
        except Exception:
            pass
        faulthandler.enable(file=fh, all_threads=True)
        faulthandler.dump_traceback_later(60, repeat=True, file=fh)

        def _heartbeat():
            local_fh = fh
            if local_fh is None:
                return
            try:
                local_fh.write(f"heartbeat {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
                local_fh.flush()
            except Exception:
                return
            try:
                root.after(1000, _heartbeat)
            except Exception:
                return

        root.after(1000, _heartbeat)
    except Exception:
        if fh:
            with contextlib.suppress(Exception):
                fh.close()
        fh = None
    return dump_path, fh


def main() -> int:
    if sys.platform == "win32":
        windll = getattr(ctypes, "windll", None)
        if windll is not None:
            try:
                windll.shcore.SetProcessDpiAwareness(2)
            except (AttributeError, OSError):
                with contextlib.suppress(AttributeError, OSError):
                    windll.user32.SetProcessDPIAware()

    logger = logging.getLogger()
    try:
        from docwen.i18n import t as gui_t
        from docwen.translation import set_translator

        set_translator(gui_t)
    except Exception:
        pass

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
            else:
                logger.error("发送文件命令失败")
        else:
            command = {"action": "activate"}
            success = FileIPC.send_command(ipc_dir, command)
            if success:
                logger.info("已发送激活窗口命令")
            else:
                logger.error("发送激活命令失败")

        return 0

    try:
        init_status = initialize_app(return_status=True)
        init_ok, init_msg = init_status if init_status else (True, None)
        if not init_ok and init_msg:
            show_error_dialog("安全检查降级", init_msg)
    except SecurityCheckFailedError as e:
        show_error_dialog("安全检查失败", str(e))
        instance_lock.release()
        return 1

    try:
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
            import time as _time

            import ttkbootstrap as tb
            from tkinterdnd2 import TkinterDnD

            _t_tk0 = _time.perf_counter()
            root = TkinterDnD.Tk()
            logger.info("⏱ TkinterDnD.Tk() 创建耗时: %.3fs", _time.perf_counter() - _t_tk0)

            _t_style0 = _time.perf_counter()
            style = tb.Style()
            logger.info("⏱ ttkbootstrap Style() 创建耗时: %.3fs", _time.perf_counter() - _t_style0)

            from docwen.gui.core.typing_utils import as_tb_window, attach_style

            attach_style(root, style)
            root.title("公文转换器")

            dump_path, dump_fh = _setup_gui_hang_diagnostics(root)
            logger.info(
                "GUI 卡死诊断日志: %s (每60秒自动dump线程堆栈；可用环境变量 DOCWEN_GUI_HANG_DIR 修改目录；自动保留最近10个且最多保留7天)",
                dump_path,
            )

            try:
                from docwen.utils.icon_utils import get_icon_path

                icon_path = get_icon_path()
                if icon_path and Path(icon_path).exists():
                    root.iconbitmap(icon_path)
            except Exception:
                pass

            from docwen.config.config_manager import config_manager
            from docwen.gui.core.window import MainWindow

            initial_file = sys.argv[1] if len(sys.argv) > 1 else None
            app = MainWindow(as_tb_window(root), config_manager=config_manager, initial_file_path=initial_file)

            ipc_dir = instance_lock.get_ipc_dir()
            file_ipc = FileIPC(ipc_dir, app.handle_ipc_command)
            file_ipc.start()

            try:
                logger.info("进入 Tk mainloop")
                root.mainloop()
                logger.info("退出 Tk mainloop")
            finally:
                with contextlib.suppress(Exception):
                    faulthandler.cancel_dump_traceback_later()
                if dump_fh:
                    with contextlib.suppress(Exception):
                        dump_fh.close()
                file_ipc.stop()
                instance_lock.release()

            return 0
        except ImportError as e:
            show_error_dialog("导入错误", f"无法导入GUI模块: {e!s}")
            return 1
        except Exception as e:
            show_error_dialog("启动错误", f"GUI启动失败: {e!s}")
            return 1
    except Exception as e:
        logger.critical(f"GUI主程序崩溃: {e!s}", exc_info=True)
        show_error_dialog("致命错误", f"程序崩溃: {e!s}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
