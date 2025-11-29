"""
策略基类模块
定义了所有转换策略的抽象基类。
"""

import abc
from typing import Dict, Any, Callable, Optional
from gongwen_converter.services.result import ConversionResult

class BaseStrategy(abc.ABC):
    """
    所有转换策略的抽象基类。
    """

    @abc.abstractmethod
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行策略的核心方法。

        Args:
            file_path (str): 要处理的输入文件路径。
            options (Optional[Dict[str, Any]]): 包含额外参数的字典，
                例如 'template_name', 'spell_check_options' 等。
            progress_callback (Optional[Callable[[str], None]]): 一个可选的回调函数，
                用于向调用方报告进度更新。回调函数接收一个字符串参数，即状态消息。

        Returns:
            ConversionResult: 包含操作结果的标准化对象。
        """
        pass
