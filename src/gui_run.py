"""
公文转换器图形用户界面启动模块

本模块是公文转换器GUI应用程序的主入口点，负责整个GUI启动流程的协调和管理。
主要功能包括：
- 初始化极简日志系统用于启动阶段错误记录
- 检查GUI所需的第三方依赖包是否可用
- 加载和应用GUI配置设置
- 初始化正式日志系统
- 创建并启动主应用程序窗口

启动流程分为多个阶段，确保在出现错误时能够提供友好的用户反馈。
"""

import os
import sys
import logging
import platform
import ctypes
from typing import Optional, Tuple

# 统一日志系统: 在所有其他导入之前进行预初始化
from gongwen_converter.utils.logging_utils import pre_init_logging, init_logging_system
pre_init_logging()

from gongwen_converter.bootstrap import initialize_app
from gongwen_converter.utils.gui_utils import show_info_dialog, show_error_dialog

def check_dependencies() -> Tuple[bool, Optional[str]]:
    """
    检查GUI所需的依赖是否安装
    
    返回:
        Tuple[bool, Optional[str]]: (是否所有依赖可用, 错误消息)
    """
    logger = logging.getLogger()
    logger.info("开始检查GUI依赖...")
    
    # 检查Tkinter
    try:
        import tkinter
        tk_version = tkinter.TkVersion
        logger.debug(f"Tkinter可用，版本: {tk_version}")
    except ImportError:
        error_msg = "Tkinter不可用，GUI无法启动。"
        logger.error(error_msg)
        return False, error_msg
    
    # 检查tkinterdnd2
    try:
        import tkinterdnd2
        logger.debug("tkinterdnd2可用。")
    except ImportError:
        error_msg = "tkinterdnd2未安装，请运行: pip install tkinterdnd2。"
        logger.error(error_msg)
        return False, error_msg
    
    # 检查ttkbootstrap
    try:
        import ttkbootstrap
        logger.debug(f"ttkbootstrap可用。")
    except ImportError:
        error_msg = "ttkbootstrap未安装，请运行: pip install ttkbootstrap。"
        logger.error(error_msg)
        return False, error_msg

    logger.info("所有GUI依赖检查通过。")
    return True, None

def load_gui_config() -> bool:
    """
    加载GUI配置（在启动主窗口前）
    
    返回:
        bool: 是否成功加载配置
    """
    logger = logging.getLogger()
    logger.info("开始加载GUI配置...")
    
    try:
        # 导入配置管理器
        from gongwen_converter.config.config_manager import config_manager
        
        # 检查GUI配置是否已加载
        gui_config = config_manager.get_gui_config_block()
        if not gui_config:
            logger.warning("GUI配置为空，使用默认设置")
            return False
        
        # 记录一些关键配置
        width = config_manager.get_center_panel_width()
        height = config_manager.get_window_height()
        theme = config_manager.get_default_theme()
        transparency = config_manager.get_transparency_value()
        
        logger.info(f"GUI配置加载成功: 尺寸={width}x{height}, 主题={theme}, 透明度={transparency}")
        return True
        
    except ImportError as e:
        logger.error(f"导入配置管理器失败: {str(e)}")
        return False
    except Exception as e:
        logger.error(f"加载GUI配置失败: {str(e)}")
        return False

def main() -> int:
    """
    GUI主入口函数
    
    返回:
        int: 退出代码 (0表示成功，非0表示错误)
    """
    # ==== 阶段0: 设置DPI感知 ====
    if platform.system() == "Windows":
        try:
            # 设置为 Per-Monitor DPI-Aware V2
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except (AttributeError, OSError):
            # 在旧版Windows上可能失败，回退到系统级DPI感知
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except (AttributeError, OSError):
                pass  # 忽略错误，程序仍可运行

    # 日志系统已在模块导入时预初始化
    logger = logging.getLogger()
    
    # ==== 阶段0.5: 单实例检查 ====
    from gongwen_converter.ipc.single_instance import SingleInstance
    from gongwen_converter.ipc.file_ipc import FileIPC
    
    instance_lock = SingleInstance("gongwen_converter")
    
    if not instance_lock.acquire():
        # 程序已在运行
        logger.info("检测到程序已在运行，将命令发送到已运行的实例")
        
        ipc_dir = instance_lock.get_ipc_dir()
        
        # 如果有命令行参数（文件路径），发送给已运行的实例
        if len(sys.argv) > 1:
            file_path = sys.argv[1]
            command = {
                'action': 'add_file',
                'file_path': file_path,
                'mode': 'single'  # 默认单文件模式
            }
            success = FileIPC.send_command(ipc_dir, command)
            if success:
                logger.info(f"已将文件发送到运行中的程序: {file_path}")
                print(f"已将文件发送到运行中的程序: {file_path}")
            else:
                logger.error("发送文件命令失败")
                print("发送文件命令失败")
        else:
            # 没有参数，发送激活命令
            command = {'action': 'activate'}
            success = FileIPC.send_command(ipc_dir, command)
            if success:
                logger.info("已发送激活窗口命令")
                print("程序已在运行，已激活窗口")
            else:
                logger.error("发送激活命令失败")
                print("发送激活命令失败")
        
        return 0
    
    # ==== 阶段1: 统一应用初始化 ====
    # 调用新的启动器来处理安全、网络隔离和路径问题
    initialize_app()
    
    try:
        # ==== 阶段2：有效期检查（仅获取状态，不显示弹窗） ====
        from gongwen_converter.security.expiration_check import get_expiration_status, ExpirationStatus
        
        # 获取有效期状态，稍后在主窗口启动后显示
        expiration_status_info = get_expiration_status()
        logger.info(f"有效期状态: {expiration_status_info}")

        # ==== 阶段3：依赖检查 ====
        logger.info("开始依赖检查...")
        deps_ok, error_msg = check_dependencies()
        if not deps_ok:
            logger.error("依赖检查失败，GUI无法启动。")
            show_error_dialog("依赖缺失", error_msg or "缺少必要的依赖包。")
            return 1
        
        # ==== 阶段4：加载GUI配置 ====
        logger.info("加载GUI配置...")
        if not load_gui_config():
            logger.warning("GUI配置加载失败，将使用默认设置。")
        
        # ==== 阶段5：正式日志系统 ====
        logger.info("初始化正式日志系统...")
        try:
            init_logging_system()
            logger.info("正式日志系统初始化成功。")
        except Exception as e:
            logger.critical(f"正式日志系统初始化失败: {e}", exc_info=True)
            # 即使失败，程序仍可依赖预初始化日志继续运行
        
        # ==== 阶段6：启动GUI ====
        logger.info("启动GUI主窗口...")
        
        # 导入并创建GUI主窗口
        try:
            import ttkbootstrap as tb
            from tkinterdnd2 import TkinterDnD
            
            # 创建支持拖拽的根窗口
            root = TkinterDnD.Tk()
            
            # 设置ttkbootstrap样式（主题将在主窗口内部根据配置设置）
            style = tb.Style()
            root.style = style
            
            # 设置窗口标题
            root.title("公文转换器")
            
            # 设置窗口图标（使用 icon_utils 模块）
            try:
                from gongwen_converter.utils.icon_utils import get_icon_path
                icon_path = get_icon_path()
                if icon_path and os.path.exists(icon_path):
                    root.iconbitmap(icon_path)
                    logger.info(f"设置窗口图标: {icon_path}")
                else:
                    logger.warning("图标文件不存在，使用默认图标")
            except ImportError as e:
                logger.warning(f"无法导入 icon_utils: {str(e)}，使用默认图标")
            except Exception as e:
                logger.warning(f"设置图标失败: {str(e)}，使用默认图标")
            
            # 在创建主窗口前，导入配置管理器
            from gongwen_converter.config.config_manager import config_manager

            # 处理命令行参数
            initial_file = sys.argv[1] if len(sys.argv) > 1 else None
            
            # 创建主窗口，并注入config_manager实例
            from gongwen_converter.gui.core.window import MainWindow
            app = MainWindow(root, config_manager=config_manager, initial_file_path=initial_file)
            
            # 延迟显示有效期提醒（在主窗口完全初始化后）
            def show_expiration_reminder():
                """延迟显示有效期提醒弹窗"""
                if expiration_status_info.status == ExpirationStatus.NEARING_EXPIRATION:
                    show_info_dialog(
                        "试用期提醒",
                        f"此软件的测试版本有效期不到3个月（剩余 {expiration_status_info.days_left} 天）。\n到期后，部分操作可能受限，请及时联系开发者。"
                    )
                elif expiration_status_info.status == ExpirationStatus.EXPIRED:
                    show_info_dialog(
                        "试用期已到期",
                        "此软件的测试版本已过期。\n请联系开发者获取正式版本。",
                        alert=True
                    )
            
            # 在主循环开始后延迟200ms显示有效期提醒
            root.after(200, show_expiration_reminder)
            
            # 启动文件 IPC 监听
            ipc_dir = instance_lock.get_ipc_dir()
            file_ipc = FileIPC(ipc_dir, app.handle_ipc_command)
            if file_ipc.start():
                logger.info("文件 IPC 已启动")
            else:
                logger.warning("文件 IPC 启动失败，将无法接收外部命令")
            
            # 运行主循环
            logger.info("GUI主循环启动")
            try:
                root.mainloop()
            finally:
                # 清理资源
                logger.info("正在清理资源...")
                file_ipc.stop()
                instance_lock.release()
            
            logger.info("GUI正常退出")
            return 0
            
        except ImportError as e:
            logger.error(f"导入GUI模块失败: {str(e)}")
            show_error_dialog("导入错误", f"无法导入GUI模块: {str(e)}")
            return 1
        except Exception as e:
            logger.critical(f"GUI启动失败: {str(e)}", exc_info=True)
            show_error_dialog("启动错误", f"GUI启动失败: {str(e)}")
            return 1
            
    except Exception as e:
        # 使用预初始化日志记录严重错误
        logger.critical(f"GUI主程序崩溃: {str(e)}", exc_info=True)
        show_error_dialog("致命错误", f"程序崩溃: {str(e)}")
        return 1

if __name__ == "__main__":
    # 运行GUI并返回退出代码
    exit_code = main()
    sys.exit(exit_code)
