"""
服务层结果对象模块
定义了所有服务层操作返回的标准数据结构。
"""

from typing import NamedTuple, Optional

class ConversionResult(NamedTuple):
    """
    一个标准化的转换结果容器。

    Attributes:
        success (bool): 操作是否成功。
        output_path (Optional[str]): 成功时生成的输出文件路径。
        message (Optional[str]): 附加的用户友好消息（成功或失败原因）。
        error (Optional[Exception]): 如果发生异常，则包含异常对象。
    """
    success: bool
    output_path: Optional[str] = None
    message: Optional[str] = None
    error: Optional[Exception] = None
