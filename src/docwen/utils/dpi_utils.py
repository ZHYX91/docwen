"""
DPI缩放管理工具模块

提供统一的高DPI屏幕适配功能，包括：
- 自动检测系统DPI
- 单例模式的缩放管理器
- 便捷的缩放混入类

DPI缩放配置说明：
==================
ENABLE_DPI_SCALING = True:  启用DPI缩放（默认）
  - 在高DPI屏幕（如200%缩放）上，UI会自动放大
  - 适合大多数用户，提供更好的显示效果

ENABLE_DPI_SCALING = False: 禁用DPI缩放
  - 在任何DPI环境下都使用原始设计尺寸
  - 在高DPI屏幕上UI会显得较小
  - 适合希望在高DPI屏幕上显示更多内容的用户
"""

import ctypes
import logging
import platform
from typing import Optional

logger = logging.getLogger(__name__)


class DPIManager:
    """
    DPI缩放管理器（单例模式）

    负责检测系统DPI并提供统一的缩放计算功能。
    全局只会创建一个实例，确保整个应用使用相同的缩放因子。
    """

    # ========== 硬编码开关 ==========
    ENABLE_DPI_SCALING: bool = True  # False = 禁用DPI缩放，使用原始大小
    # ================================

    _instance: Optional["DPIManager"] = None
    _initialized: bool = False

    def __new__(cls):
        """单例模式：确保全局只有一个实例"""
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self):
        """初始化DPI管理器"""
        if not DPIManager._initialized:
            self.scaling_factor: float = 1.0
            self._root_window = None
            DPIManager._initialized = True

    def initialize(self, root_window=None) -> None:
        """
        初始化并检测DPI缩放因子

        参数:
            root_window: Tkinter根窗口实例（可选）
        """
        self._root_window = root_window
        self._detect_scaling_factor()

    def _detect_scaling_factor(self) -> None:
        """检测系统DPI并计算缩放因子"""
        if not self.ENABLE_DPI_SCALING:
            self.scaling_factor = 1.0
            logger.info("DPI缩放已禁用（硬编码开关），缩放因子固定为: 1.0")
            logger.info("所有UI元素将使用原始设计尺寸显示")
            return

        try:
            if platform.system() == "Windows":
                self._detect_windows_dpi()
            else:
                self._detect_tkinter_dpi()
        except Exception as e:
            logger.error(f"检测DPI失败: {e}，使用默认缩放因子 1.0")
            self.scaling_factor = 1.0

    def _detect_windows_dpi(self) -> None:
        """检测Windows系统的DPI"""
        try:
            # 尝试通过窗口句柄获取DPI（最准确）
            if self._root_window:
                self._root_window.update_idletasks()
                hwnd = ctypes.windll.user32.GetParent(self._root_window.winfo_id())
                dpi = ctypes.windll.user32.GetDpiForWindow(hwnd)
                self.scaling_factor = dpi / 96.0
                logger.info(f"Windows DPI检测: {dpi}, 缩放因子: {self.scaling_factor:.2f}")
                return
        except Exception as e:
            logger.warning(f"通过窗口句柄获取DPI失败: {e}，尝试备用方案")

        # 备用方案：使用Tkinter的缩放因子
        try:
            if self._root_window:
                self.scaling_factor = self._root_window.tk.call("tk", "scaling")
                logger.info(f"使用Tkinter缩放因子: {self.scaling_factor:.2f}")
            else:
                self.scaling_factor = 1.0
                logger.warning("无根窗口引用，使用默认缩放因子: 1.0")
        except Exception as e:
            logger.error(f"获取Tkinter缩放因子失败: {e}")
            self.scaling_factor = 1.0

    def _detect_tkinter_dpi(self) -> None:
        """检测非Windows系统的DPI（通过Tkinter）"""
        try:
            if self._root_window:
                self.scaling_factor = self._root_window.tk.call("tk", "scaling")
                logger.info(f"系统DPI检测（Tkinter）: 缩放因子 {self.scaling_factor:.2f}")
            else:
                self.scaling_factor = 1.0
                logger.warning("无根窗口引用，使用默认缩放因子: 1.0")
        except Exception as e:
            logger.error(f"获取系统缩放因子失败: {e}")
            self.scaling_factor = 1.0

    def scale(self, value: int) -> int:
        """
        根据缩放因子调整数值

        参数:
            value: 原始像素值

        返回:
            缩放后的像素值
        """
        if not self.ENABLE_DPI_SCALING:
            return value
        return int(value * self.scaling_factor)

    def unscale(self, value: int) -> int:
        if not self.ENABLE_DPI_SCALING:
            return value
        if self.scaling_factor == 0:
            return value
        return round(value / self.scaling_factor)

    def get_scaling_factor(self) -> float:
        """获取当前缩放因子"""
        return self.scaling_factor

    def is_scaling_enabled(self) -> bool:
        """检查DPI缩放是否启用"""
        return self.ENABLE_DPI_SCALING


class ScalableMixin:
    """
    可缩放混入类

    任何需要DPI缩放的类都可以继承此混入类，自动获得scale()方法。
    无需手动传递scaling_factor参数。

    使用示例:
        class MyDialog(tk.Toplevel, ScalableMixin):
            def __init__(self, parent):
                super().__init__(parent)
                width = self.scale(800)  # 自动缩放
                height = self.scale(600)
                self.geometry(f"{width}x{height}")
    """

    def scale(self, value: int) -> int:
        """
        根据全局DPI缩放因子调整数值

        参数:
            value: 原始像素值

        返回:
            缩放后的像素值
        """
        return dpi_manager.scale(value)

    def get_scaling_factor(self) -> float:
        """获取当前缩放因子"""
        return dpi_manager.get_scaling_factor()

    def unscale(self, value: int) -> int:
        return dpi_manager.unscale(value)


# 创建全局单例实例
dpi_manager = DPIManager()


def initialize_dpi_manager(
    root_window=None,
    *,
    scaling_enabled: bool | None = None,
    forced_scale: float | None = None,
) -> None:
    """
    初始化全局DPI管理器

    应该在创建主窗口后立即调用此函数。

    参数:
        root_window: Tkinter根窗口实例
        scaling_enabled: 是否启用DPI缩放；None 表示沿用默认行为
        forced_scale: 强制指定缩放系数；None/0 表示自动检测
    """
    if scaling_enabled is not None:
        DPIManager.ENABLE_DPI_SCALING = bool(scaling_enabled)
    dpi_manager.initialize(root_window)

    if forced_scale is not None:
        try:
            forced_scale_value = float(forced_scale)
        except (TypeError, ValueError):
            forced_scale_value = 0.0

        if forced_scale_value and DPIManager.ENABLE_DPI_SCALING:
            if 0.5 <= forced_scale_value <= 4.0:
                dpi_manager.scaling_factor = forced_scale_value
                logger.info(f"DPI缩放系数已强制覆盖为: {forced_scale_value:.2f}")
            else:
                logger.warning(f"强制缩放系数超出合理范围(0.5~4.0): {forced_scale_value}，将使用自动检测值")

    logger.info(f"DPI管理器初始化完成，缩放因子: {dpi_manager.get_scaling_factor():.2f}")


def scale(value: int) -> int:
    """
    便捷函数：缩放数值

    参数:
        value: 原始像素值

    返回:
        缩放后的像素值
    """
    return dpi_manager.scale(value)


def unscale(value: int) -> int:
    return dpi_manager.unscale(value)


def get_scaling_factor() -> float:
    """便捷函数：获取当前缩放因子"""
    return dpi_manager.get_scaling_factor()


def is_scaling_enabled() -> bool:
    """便捷函数：检查DPI缩放是否启用"""
    return dpi_manager.is_scaling_enabled()
