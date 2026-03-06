"""
选项联动处理器模块

提供可复用的导出选项联动处理逻辑，处理"提取图片"和"OCR"选项的联动关系。

主要功能：
- 勾选OCR时，自动勾选"提取图片"
- 取消"提取图片"时，自动取消OCR
- 防止递归触发的状态管理

依赖：
- tkinter: 提供 BooleanVar 变量

使用示例：
    from docwen.gui.components.common import ExportOptionHandler

    # 创建处理器
    handler = ExportOptionHandler(image_var, ocr_var)

    # 绑定到复选框的command
    checkbox.config(command=handler.on_option_changed)
"""

import logging
import tkinter as tk
from collections.abc import Callable
from typing import Any

logger = logging.getLogger(__name__)


class ExportOptionHandler:
    """
    导出选项联动处理器

    处理"提取图片"和"OCR识别"两个选项的联动逻辑：
    - 勾选OCR时，自动勾选"提取图片"（因为OCR依赖图片提取）
    - 取消"提取图片"时，自动取消OCR（因为没有图片就无法OCR）

    使用状态变化检测避免死循环联动。

    属性：
        image_var: 提取图片选项的 BooleanVar
        ocr_var: OCR识别选项的 BooleanVar
        on_state_changed: 状态变化后的回调函数（可选）
    """

    def __init__(
        self, image_var: tk.BooleanVar, ocr_var: tk.BooleanVar, on_state_changed: Callable[[], None] | None = None
    ):
        """
        初始化选项联动处理器

        参数：
            image_var: 提取图片选项的 BooleanVar
            ocr_var: OCR识别选项的 BooleanVar
            on_state_changed: 状态变化后的回调函数，用于更新UI或执行其他逻辑
        """
        self.image_var = image_var
        self.ocr_var = ocr_var
        self.on_state_changed = on_state_changed

        # 状态记录（用于检测变化）
        self._last_image_state = image_var.get()
        self._last_ocr_state = ocr_var.get()

        logger.debug(f"初始化选项联动处理器: 提取图片={self._last_image_state}, OCR={self._last_ocr_state}")

    def on_option_changed(self) -> None:
        """处理选项变更事件（提取图片和OCR独立控制，无联动）"""
        # 更新状态记录
        self._last_image_state = self.image_var.get()
        self._last_ocr_state = self.ocr_var.get()

        # 调用状态变化回调
        if self.on_state_changed:
            self.on_state_changed()

    def reset_state(self) -> None:
        """
        重置状态记录

        当选项被外部代码修改时调用，确保状态记录与实际值同步。
        """
        self._last_image_state = self.image_var.get()
        self._last_ocr_state = self.ocr_var.get()
        logger.debug(f"重置状态记录: 提取图片={self._last_image_state}, OCR={self._last_ocr_state}")

    def get_options(self) -> dict:
        """
        获取当前选项值

        返回：
            dict: 包含 'extract_image' 和 'extract_ocr' 的字典
        """
        return {"extract_image": self.image_var.get(), "extract_ocr": self.ocr_var.get()}

    def is_any_selected(self) -> bool:
        """
        检查是否至少选择了一个选项

        返回：
            bool: 如果至少有一个选项被选中则返回 True
        """
        return self.image_var.get() or self.ocr_var.get()


class NumberingOptionHandler:
    """
    序号选项联动处理器

    处理"清除序号"和"添加序号"选项的联动逻辑，
    以及"添加序号"复选框与序号方案下拉框的联动。

    属性：
        remove_var: 清除序号选项的 BooleanVar
        add_var: 添加序号选项的 BooleanVar
        scheme_combo: 序号方案下拉框（可选）
    """

    def __init__(
        self,
        remove_var: tk.BooleanVar,
        add_var: tk.BooleanVar,
        scheme_combo: Any | None = None,
        on_state_changed: Callable[[], None] | None = None,
    ):
        """
        初始化序号选项联动处理器

        参数：
            remove_var: 清除序号选项的 BooleanVar
            add_var: 添加序号选项的 BooleanVar
            scheme_combo: 序号方案下拉框组件
            on_state_changed: 状态变化后的回调函数
        """
        self.remove_var = remove_var
        self.add_var = add_var
        self.scheme_combo = scheme_combo
        self.on_state_changed = on_state_changed

        logger.debug(f"初始化序号选项联动处理器: 清除序号={remove_var.get()}, 添加序号={add_var.get()}")

    def on_add_numbering_toggle(self) -> None:
        """
        处理"添加序号"复选框切换事件

        控制序号方案下拉框的启用/禁用状态：
        - 勾选"添加序号"时：启用下拉框
        - 不勾选时：禁用下拉框（灰色）
        """
        if self.scheme_combo is None:
            return

        if self.add_var.get():
            self.scheme_combo.config(state="readonly")
            logger.debug("添加序号已启用，序号方案下拉框可选")
        else:
            self.scheme_combo.config(state="disabled")
            logger.debug("添加序号已禁用，序号方案下拉框灰色")

        # 调用状态变化回调
        if self.on_state_changed:
            self.on_state_changed()

    def get_options(self, scheme_name_to_id: dict[str, str] | None = None) -> dict[str, object]:
        """
        获取当前选项值

        参数：
            scheme_name_to_id: 方案名称到ID的映射字典（可选）

        返回：
            dict: 包含序号相关选项的字典
        """
        options: dict[str, object] = {"remove_numbering": self.remove_var.get(), "add_numbering": self.add_var.get()}

        # 如果有下拉框和映射字典，添加方案ID
        if self.scheme_combo and scheme_name_to_id:
            try:
                scheme_name = self.scheme_combo.get()
                options["numbering_scheme"] = scheme_name_to_id.get(scheme_name, "gongwen_standard")
            except Exception as e:
                logger.warning(f"获取序号方案失败: {e}")
                options["numbering_scheme"] = "gongwen_standard"

        return options
