"""
全局主题管理器模块
提供统一的主题管理，通过主窗口全局应用主题
"""

import logging
import tkinter as tk
from typing import Optional
import ttkbootstrap as tb

logger = logging.getLogger(__name__)


class ThemeManager:
    """
    全局主题管理器（单例模式）
    通过主窗口应用主题，所有子窗口自动继承
    """
    
    _instance: Optional['ThemeManager'] = None
    
    def __init__(self):
        """私有构造函数，使用 get_instance() 获取实例"""
        if ThemeManager._instance is not None:
            raise RuntimeError("ThemeManager 是单例类，请使用 get_instance() 方法获取实例")
        
        self._root: Optional[tk.Tk] = None  # 主窗口引用
        self._current_theme: Optional[str] = None  # 当前主题名称
        
        logger.debug("ThemeManager 初始化完成")
    
    @classmethod
    def get_instance(cls) -> 'ThemeManager':
        """
        获取 ThemeManager 单例实例
        
        返回:
            ThemeManager: 单例实例
        """
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance
    
    @classmethod
    def reset_instance(cls):
        """重置单例实例（主要用于测试）"""
        cls._instance = None
        logger.debug("ThemeManager 实例已重置")
    
    def initialize(self, root: tb.Window, initial_theme: str):
        """
        初始化主题管理器
        
        参数:
            root: 主窗口对象
            initial_theme: 初始主题名称
        """
        self._root = root
        self._current_theme = initial_theme
        self._apply_to_window(root, initial_theme)
        logger.info(f"ThemeManager 初始化完成，当前主题: {initial_theme}")
    
    def apply_theme(self, theme_name: str, preview_only: bool = False) -> bool:
        """
        应用主题到主窗口（所有子窗口自动继承）
        
        参数:
            theme_name: 主题名称
            preview_only: 是否仅为预览（不更新 _current_theme）
            
        返回:
            bool: 是否成功应用主题
        """
        logger.info(f"========== 开始应用主题 ==========")
        logger.info(f"目标主题: {theme_name}")
        logger.info(f"预览模式: {preview_only}")
        logger.info(f"当前主题: {self._current_theme}")
        
        try:
            # 检查主窗口是否存在
            logger.debug(f"检查主窗口状态...")
            if self._root is None:
                logger.error("❌ 主窗口引用为 None")
                return False
            
            try:
                window_exists = self._root.winfo_exists()
                logger.debug(f"主窗口存在性检查: {window_exists}")
                if not window_exists:
                    logger.error("❌ 主窗口已被销毁")
                    return False
            except tk.TclError as e:
                logger.error(f"❌ 主窗口存在性检查失败: {e}")
                return False
            
            # 检查主题是否可用
            logger.debug(f"获取可用主题列表...")
            try:
                available_themes = self._root.style.theme_names()
                logger.debug(f"可用主题列表: {available_themes}")
                
                if theme_name not in available_themes:
                    logger.error(f"❌ 主题 '{theme_name}' 不在可用列表中")
                    logger.error(f"可用主题: {', '.join(available_themes)}")
                    return False
                else:
                    logger.debug(f"✓ 主题 '{theme_name}' 存在于可用列表中")
            except Exception as e:
                logger.error(f"❌ 获取主题列表失败: {e}")
                return False
            
            # 先处理所有待处理的事件（如组件销毁）
            logger.debug(f"处理待处理的UI事件...")
            try:
                self._root.update_idletasks()
                logger.debug(f"✓ 待处理事件已处理完成")
            except tk.TclError as e:
                logger.debug(f"处理待处理事件时出错（可能窗口已销毁）: {e}")
            
            # 只对主窗口应用主题 - 所有子窗口会自动继承
            logger.info(f"开始应用主题到主窗口...")
            success = self._apply_to_window(self._root, theme_name)
            
            if success:
                # 如果不是预览模式，更新当前主题
                if not preview_only:
                    old_theme = self._current_theme
                    self._current_theme = theme_name
                    logger.info(f"✓ 主题已更新: {old_theme} -> {theme_name}")
                else:
                    logger.debug(f"✓ 主题预览应用完成: {theme_name}")
                
                logger.info(f"✓ 主题 '{theme_name}' 已成功应用到主窗口（所有子窗口自动继承）")
                logger.info(f"========== 主题应用成功 ==========")
            else:
                logger.error(f"❌ 应用主题 '{theme_name}' 失败")
                logger.info(f"========== 主题应用失败 ==========")
            
            return success
            
        except Exception as e:
            logger.error(f"❌ 应用主题时发生异常: {type(e).__name__}: {str(e)}")
            logger.exception("异常详情:")
            logger.info(f"========== 主题应用异常 ==========")
            return False
    
    def _apply_to_window(self, window: tk.Misc, theme_name: str) -> bool:
        """
        应用主题到单个窗口
        
        参数:
            window: 窗口对象
            theme_name: 主题名称
            
        返回:
            bool: 是否成功应用
        """
        try:
            # 检查窗口是否存在
            logger.debug(f"  >> 检查窗口对象: {window}")
            try:
                window_exists = window.winfo_exists()
                logger.debug(f"  >> 窗口存在性: {window_exists}")
                if not window_exists:
                    logger.debug(f"  >> 窗口不存在，跳过")
                    return False
            except tk.TclError as e:
                logger.debug(f"  >> 窗口存在性检查失败 (已销毁): {e}")
                return False
            
            window_class = window.winfo_class()
            window_path = str(window)
            logger.debug(f"  >> 窗口类型: {window_class}")
            logger.debug(f"  >> 窗口路径: {window_path}")
            
            # 获取 style 对象并应用主题
            # ttkbootstrap 的窗口都应该有 style 属性
            has_style = hasattr(window, 'style')
            logger.debug(f"  >> 是否有 style 属性: {has_style}")
            
            if has_style:
                logger.debug(f"  >> 窗口 {window_class} 有 style 属性，直接应用")
                try:
                    # 获取当前主题（用于日志）
                    try:
                        current_theme = window.style.theme_use()
                        logger.debug(f"  >> 当前主题: {current_theme}")
                    except:
                        logger.debug(f"  >> 无法获取当前主题")
                    
                    # 应用新主题
                    logger.debug(f"  >> 调用 window.style.theme_use('{theme_name}')...")
                    
                    # 双重保护：即使 theme_use 对某些组件失败，也要确保主窗口的主题被应用
                    theme_applied = False
                    try:
                        window.style.theme_use(theme_name)
                        theme_applied = True
                        logger.debug(f"  >> ✓ theme_use() 调用成功")
                    except tk.TclError as theme_error:
                        error_msg = str(theme_error)
                        logger.debug(f"  >> theme_use() TclError: {error_msg}")
                        
                        # 检查是否是子组件销毁导致的错误
                        if "bad window path name" in error_msg:
                            # 提取出错的组件路径
                            logger.debug(f"  >> 检测到子组件销毁错误，这是正常的")
                            # 即使有子组件错误，主题可能已部分应用
                            # 我们认为这是成功的
                            theme_applied = True
                        else:
                            # 其他类型的错误才认为是失败
                            logger.warning(f"  >> ✗ theme_use() 失败: {error_msg}")
                            raise
                    
                    if not theme_applied:
                        logger.debug(f"  >> ✗ 主题未成功应用")
                        return False
                    
                    # 更新界面
                    logger.debug(f"  >> 调用 window.update_idletasks()...")
                    try:
                        window.update_idletasks()
                        logger.debug(f"  >> ✓ update_idletasks() 完成")
                    except tk.TclError as update_error:
                        # update_idletasks 的错误不影响主题应用成功
                        logger.debug(f"  >> update_idletasks() 出错（不影响主题应用）: {update_error}")
                    
                    logger.debug(f"  >> ✓ 成功应用主题到 {window_class}")
                    return True
                    
                except tk.TclError as e:
                    error_msg = str(e)
                    logger.debug(f"  >> TclError: {error_msg}")
                    if "bad window path name" in error_msg or "has been destroyed" in error_msg:
                        logger.debug(f"  >> 窗口 {window_class} 已销毁，跳过应用主题")
                    else:
                        logger.warning(f"  >> ✗ 窗口 {window_class} 应用主题失败 (TclError): {error_msg}")
                    return False
                except Exception as e:
                    logger.warning(f"  >> ✗ 窗口 {window_class} 应用主题失败: {type(e).__name__}: {e}")
                    logger.exception("  >> 异常详情:")
                    return False
            else:
                # 回退：使用根窗口的 style
                logger.debug(f"  >> 窗口 {window_class} 无 style 属性，尝试使用根窗口 style")
                if self._root and hasattr(self._root, 'style'):
                    logger.debug(f"  >> 根窗口有 style 属性")
                    try:
                        logger.debug(f"  >> 调用根窗口的 style.theme_use('{theme_name}')...")
                        self._root.style.theme_use(theme_name)
                        
                        logger.debug(f"  >> 调用 window.update_idletasks()...")
                        window.update_idletasks()
                        
                        logger.debug(f"  >> ✓ 通过根窗口成功应用主题到 {window_class}")
                        return True
                    except tk.TclError as e:
                        error_msg = str(e)
                        logger.debug(f"  >> TclError: {error_msg}")
                        if "bad window path name" in error_msg or "has been destroyed" in error_msg:
                            logger.debug(f"  >> 窗口 {window_class} 已销毁，跳过应用主题")
                        else:
                            logger.warning(f"  >> ✗ 根窗口应用主题失败 (TclError): {error_msg}")
                        return False
                    except Exception as e:
                        logger.warning(f"  >> ✗ 根窗口应用主题失败: {type(e).__name__}: {e}")
                        logger.exception("  >> 异常详情:")
                        return False
                else:
                    logger.warning(f"  >> ✗ 无法为窗口 {window_class} 应用主题：根窗口无 style")
                    return False
            
        except tk.TclError as e:
            error_msg = str(e)
            logger.debug(f"  >> 捕获 TclError: {error_msg}")
            if "bad window path name" not in error_msg and "has been destroyed" not in error_msg:
                logger.warning(f"  >> 应用主题到窗口失败 (TclError): {error_msg}")
            else:
                logger.debug(f"  >> 窗口已销毁，跳过应用主题")
            return False
        except Exception as e:
            logger.warning(f"  >> 应用主题到窗口时发生异常: {type(e).__name__}: {str(e)}")
            logger.exception("  >> 异常详情:")
            return False
    
    def get_current_theme(self) -> Optional[str]:
        """
        获取当前主题名称
        
        返回:
            str: 当前主题名称，如果未设置则返回 None
        """
        return self._current_theme
    
    def get_available_themes(self) -> list[str]:
        """
        获取所有可用的主题列表
        
        返回:
            list[str]: 可用主题名称列表
        """
        if self._root is None:
            logger.warning("主窗口未初始化，无法获取主题列表")
            return []
        
        try:
            return list(self._root.style.theme_names())
        except Exception as e:
            logger.error(f"获取主题列表失败: {str(e)}")
            return []


# 便捷函数，用于快速获取 ThemeManager 实例
def get_theme_manager() -> ThemeManager:
    """
    获取全局主题管理器实例
    
    返回:
        ThemeManager: 主题管理器实例
    """
    return ThemeManager.get_instance()


# 模块测试代码
if __name__ == "__main__":
    # 配置日志
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    # 创建测试窗口
    root = tb.Window(title="ThemeManager 测试", themename="morph")
    
    # 初始化 ThemeManager
    theme_mgr = get_theme_manager()
    theme_mgr.initialize(root, "morph")
    
    # 创建测试对话框（无需手动注册）
    dialog = tb.Toplevel(root, title="测试对话框")
    dialog.geometry("300x200")
    
    # 创建按钮测试主题切换
    def switch_theme():
        current = theme_mgr.get_current_theme()
        themes = theme_mgr.get_available_themes()
        current_index = themes.index(current) if current in themes else 0
        next_index = (current_index + 1) % len(themes)
        next_theme = themes[next_index]
        
        logger.info(f"切换主题: {current} -> {next_theme}")
        theme_mgr.apply_theme(next_theme)
    
    button = tb.Button(root, text="切换主题", command=switch_theme)
    button.pack(pady=20)
    
    label = tb.Label(root, text="主题会自动应用到所有窗口")
    label.pack(pady=10)
    
    logger.info("ThemeManager 测试启动")
    root.mainloop()
