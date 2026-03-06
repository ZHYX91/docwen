"""
图标工具模块
用于处理图标相关的功能
提供获取应用程序图标路径的功能，并管理UI图标的加载与缓存
"""

import logging
import sys
from pathlib import Path
from tkinter import PhotoImage
from typing import Any

# 配置日志
logger = logging.getLogger(__name__)

# 全局图标缓存
_icon_cache: dict[str, Any] = {}


def get_asset_path(asset_name: str) -> str | None:
    """
    获取assets目录下指定资源文件的完整路径

    参数:
        asset_name: 资源文件名 (例如 "icon.ico", "settings_icon.png")

    返回:
        str: 资源文件的绝对路径，如果找不到则返回None
    """
    logger.debug(f"开始查找资源文件: {asset_name}")
    try:
        from .path_utils import get_project_root

        project_root = get_project_root()

        asset_path = Path(project_root) / "assets" / asset_name

        if asset_path.exists():
            logger.debug(f"找到资源文件: {asset_path}")
            return str(asset_path)
        else:
            logger.warning(f"资源文件不存在: {asset_path}")
            return None

    except ImportError as e:
        logger.error(f"导入路径工具失败: {e!s}")
        # 回退逻辑
        if getattr(sys, "frozen", False):
            base_dir = Path(getattr(sys, "_MEIPASS", ""))
        else:
            base_dir = Path(__file__).resolve().parents[3]

        asset_path = base_dir / "assets" / asset_name
        if asset_path.exists():
            logger.info(f"使用回退方法找到资源: {asset_path}")
            return str(asset_path)

    logger.error(f"最终未找到资源文件: {asset_name}")
    return None


def get_icon_path() -> str | None:
    """
    获取主应用程序图标 (.ico) 的路径
    """
    return get_asset_path("icon.ico")


class IconManager:
    """
    图标管理器类 - 简化版
    提供全局图标管理，一次设置全局生效
    """

    _icon_path = None  # 图标路径缓存
    _initialized = False  # 初始化标记

    @classmethod
    def initialize(cls, root_window):
        """
        应用启动时调用一次，设置全局默认图标
        之后所有窗口和对话框都自动继承

        参数:
            root_window: 主窗口对象

        返回:
            bool: 是否成功设置图标
        """
        if cls._initialized:
            logger.debug("图标管理器已初始化，跳过重复初始化")
            return True

        icon_path = get_asset_path("icon.ico")
        if icon_path and Path(icon_path).exists():
            try:
                # 设置应用级默认图标，所有后续窗口自动继承
                root_window.iconbitmap(default=icon_path)
                cls._icon_path = icon_path
                cls._initialized = True
                logger.info(f"已设置应用默认图标: {icon_path}")
                return True
            except Exception as e:
                logger.error(f"设置默认图标失败: {e}")
                return False
        else:
            logger.warning("未找到图标文件，无法设置默认图标")
            return False

    @classmethod
    def get_icon_path(cls):
        """
        获取图标路径（带缓存）

        返回:
            str: 图标文件路径，如果找不到则返回None
        """
        if cls._icon_path is None:
            cls._icon_path = get_asset_path("icon.ico")
        return cls._icon_path


def load_image_icon(icon_name: str, master: Any, size: tuple[int, int] | None = None) -> Any | None:
    """
    加载PNG图标，并应用缓存机制与缩放功能

    参数:
        icon_name: 图标文件名 (例如 "settings.png")
        master: Tkinter父组件，PhotoImage需要它来初始化
        size: 可选的元组 (width, height) 用于缩放图标

    返回:
        PhotoImage: 加载并可能缩放后的图像对象，如果失败则返回None
    """
    # 使用尺寸信息创建唯一的缓存键
    cache_key = f"{icon_name}_{size}" if size else icon_name

    if cache_key in _icon_cache:
        logger.debug(f"从缓存加载图标: {cache_key}")
        return _icon_cache[cache_key]

    icon_path = get_asset_path(icon_name)
    if icon_path:
        try:
            # 尝试使用PIL进行高质量缩放
            return _load_image_with_pil(icon_path, master, size, cache_key)
        except ImportError:
            # PIL不可用，回退到原生方法
            logger.warning("PIL不可用，使用原生方法加载图标")
            return _load_image_native(icon_path, master, size, cache_key)
        except Exception as e:
            logger.error(f"使用PIL加载图标失败: {icon_path}, 错误: {e}")
            # 回退到原生方法
            return _load_image_native(icon_path, master, size, cache_key)
    return None


def _load_image_with_pil(icon_path: str, master: Any, size: tuple[int, int] | None, cache_key: str) -> Any | None:
    """
    使用PIL库加载和缩放图像（高质量）
    """
    from PIL import Image, ImageTk

    img = Image.open(icon_path)

    # 如果指定了目标尺寸，则进行缩放
    if size:
        # 使用LANCZOS重采样算法进行高质量缩放
        img = img.resize(size, Image.Resampling.LANCZOS)

    # 转换为Tkinter PhotoImage
    image = ImageTk.PhotoImage(master=master, image=img)

    # 缓存结果
    _icon_cache[cache_key] = image
    logger.info(f"使用PIL成功加载并缓存图标: {cache_key}")
    return image


def _load_image_native(icon_path: str, master: Any, size: tuple[int, int] | None, cache_key: str) -> PhotoImage | None:
    """
    使用原生Tkinter方法加载和缩放图像（兼容性回退）
    """
    try:
        image = PhotoImage(master=master, file=icon_path)

        # 如果指定了尺寸，则进行缩放
        if size:
            # 计算缩放比例
            width_ratio = image.width() // size[0]
            height_ratio = image.height() // size[1]

            if width_ratio > 0 and height_ratio > 0:
                # 需要缩小 - 使用subsample
                image = image.subsample(width_ratio, height_ratio)
            else:
                # 需要放大 - 使用zoom
                zoom_factor = max(1, int(1 / min(width_ratio, height_ratio)))
                image = image.zoom(zoom_factor, zoom_factor)

        _icon_cache[cache_key] = image
        logger.info(f"使用原生方法成功加载并缓存图标: {cache_key}")
        return image
    except Exception as e:
        logger.error(f"加载或缩放图标失败: {icon_path}, 错误: {e}")
        return None
