"""
转换上下文模块

负责管理 DOCX→MD 转换过程中的状态，包括：
- 代码块状态（in_code_block, code_block_lines, code_block_list_context）
- 列表编号状态（list_counters，委托给 ListCounterManager）
- 边框组状态（prev_border_info, prev_had_border，委托给 BorderGroupTracker）

设计说明：
    `_generate_markdown_content_simple()` 函数中原有6个状态变量相互依赖，
    拆分到独立模块后需要一个统一的上下文对象来传递和管理这些状态。
"""

import logging

logger = logging.getLogger(__name__)


class ConversionContext:
    """
    转换上下文管理器

    统一管理转换过程中的所有状态，避免在函数间传递大量参数。

    使用方式：
        ctx = ConversionContext(config_manager)

        # 代码块处理
        if is_code_block:
            ctx.start_code_block(list_context)
            ctx.add_code_line(line)
        else:
            if ctx.in_code_block:
                code_block = ctx.end_code_block()
                output(code_block)

        # 列表计数
        counter = ctx.list_counter_manager.increment(num_id, level)

        # 边框组处理
        separators = ctx.border_tracker.process_paragraph(para, config_manager)
    """

    def __init__(self, config_manager, list_indent_spaces: int = 4):
        """
        初始化转换上下文

        参数:
            config_manager: 配置管理器实例
            list_indent_spaces: 列表缩进空格数（默认4）
        """
        # 配置
        self.config_manager = config_manager
        self.list_indent_spaces = list_indent_spaces

        # 代码块状态
        self.in_code_block = False
        self.code_block_lines = []
        self.code_block_list_context = None

        # 列表计数器（延迟初始化）
        self._list_counter_manager = None

        # 边框组跟踪器（延迟初始化）
        self._border_tracker = None

        logger.debug(f"ConversionContext 初始化: list_indent_spaces={list_indent_spaces}")

    @property
    def list_counter_manager(self):
        """
        延迟初始化列表计数器管理器

        返回:
            ListCounterManager: 列表计数器管理器实例
        """
        if self._list_counter_manager is None:
            from .list_processor import ListCounterManager

            self._list_counter_manager = ListCounterManager()
            logger.debug("ListCounterManager 延迟初始化完成")
        return self._list_counter_manager

    @property
    def border_tracker(self):
        """
        延迟初始化边框组跟踪器

        返回:
            BorderGroupTracker: 边框组跟踪器实例
        """
        if self._border_tracker is None:
            from .break_processor import BorderGroupTracker

            self._border_tracker = BorderGroupTracker()
            logger.debug("BorderGroupTracker 延迟初始化完成")
        return self._border_tracker

    def start_code_block(self, list_context: tuple | None = None):
        """
        开始代码块

        参数:
            list_context: 代码块开始时的列表上下文 (numId, level)

        注意:
            代码块需要记录开始时的列表上下文，用于结束时添加正确的缩进。
        """
        if self.in_code_block:
            logger.warning("尝试开始代码块，但已经在代码块中")
            return

        self.in_code_block = True
        self.code_block_list_context = list_context

        if list_context:
            logger.debug(f"开始代码块，列表上下文: numId={list_context[0]}, level={list_context[1]}")
        else:
            logger.debug("开始代码块，无列表上下文")

    def add_code_line(self, line: str):
        """
        添加代码行

        参数:
            line: 代码行内容（原始文本，不做格式转换）
        """
        if not self.in_code_block:
            logger.warning("尝试添加代码行，但不在代码块中")
            return

        self.code_block_lines.append(line)

    def end_code_block(self) -> str:
        """
        结束代码块，返回格式化的代码块内容

        返回:
            str: 格式化的 Markdown 代码块（含缩进）
                 如果代码块为空，返回空字符串

        注意:
            此方法会自动重置代码块状态。
        """
        if not self.in_code_block:
            logger.warning("尝试结束代码块，但不在代码块中")
            return ""

        if not self.code_block_lines:
            logger.debug("代码块为空，跳过输出")
            self._reset_code_block_state()
            return ""

        # 根据列表上下文添加缩进
        if self.code_block_list_context:
            _, level = self.code_block_list_context
            indent = " " * self.list_indent_spaces * (level + 1)
            indented_lines = [indent + line for line in self.code_block_lines]
            result = indent + "```\n" + "\n".join(indented_lines) + "\n" + indent + "```"
            logger.debug(f"代码块结束，在列表上下文中，level={level}，共 {len(self.code_block_lines)} 行")
        else:
            result = "```\n" + "\n".join(self.code_block_lines) + "\n```"
            logger.debug(f"代码块结束，共 {len(self.code_block_lines)} 行")

        self._reset_code_block_state()
        return result

    def _reset_code_block_state(self):
        """重置代码块状态"""
        self.in_code_block = False
        self.code_block_lines = []
        self.code_block_list_context = None

    def reset(self):
        """
        重置所有状态（用于处理新文档）

        注意:
            当处理新文档时应调用此方法，确保状态不会相互干扰。
        """
        self._reset_code_block_state()

        if self._list_counter_manager:
            self._list_counter_manager.reset_all()

        if self._border_tracker:
            self._border_tracker.reset()

        logger.debug("ConversionContext 状态已重置")

    def __str__(self):
        """调试用字符串表示"""
        return (
            f"ConversionContext("
            f"in_code_block={self.in_code_block}, "
            f"code_lines={len(self.code_block_lines)}, "
            f"list_context={self.code_block_list_context})"
        )
