"""
版式文件策略子包

处理PDF、CAJ、OFD、XPS等版式文件的转换。

子模块：
- base: 转文档策略基类（模板方法模式）
- utils: 工具函数（预处理、中间文件管理）
- to_document: 版式转文档（DOCX/DOC/ODT/RTF/TXT）
- to_markdown: 版式转Markdown
- to_image: 版式转图片（PNG/JPG/TIF）
- to_pdf: 版式转PDF（OFD/XPS/CAJ → PDF）
- operations: PDF操作（合并/拆分）

使用方式：
    from docwen.services.strategies.layout import (
        LayoutToDocxStrategy,
        LayoutToMarkdownStrategy,
        MergePdfsStrategy,
    )
"""

# 导入基类
from .base import LayoutToDocumentBaseStrategy

# 导入工具函数
from .utils import preprocess_layout_file, should_keep_intermediates

# 导入所有策略类（触发装饰器注册）
from .to_document import (
    LayoutToTxtStrategy,
    LayoutToDocxStrategy,
    LayoutToDocStrategy,
    LayoutToOdtStrategy,
    LayoutToRtfStrategy,
)

from .to_markdown import (
    LayoutToMarkdownStrategy,
    LayoutToMarkdownPymupdf4llmStrategy,  # 兼容别名
)

from .to_image import (
    LayoutToPngStrategy,
    LayoutToJpgStrategy,
    LayoutToTifStrategy,
)

from .to_pdf import (
    LayoutToPdfStrategy,
    OfdToPdfStrategy,
    XpsToPdfStrategy,
)

from .operations import (
    MergePdfsStrategy,
    SplitPdfStrategy,
)

__all__ = [
    # 基类
    'LayoutToDocumentBaseStrategy',
    # 工具函数
    'preprocess_layout_file',
    'should_keep_intermediates',
    # 转文档策略
    'LayoutToTxtStrategy',
    'LayoutToDocxStrategy',
    'LayoutToDocStrategy',
    'LayoutToOdtStrategy',
    'LayoutToRtfStrategy',
    # 转Markdown策略
    'LayoutToMarkdownStrategy',
    'LayoutToMarkdownPymupdf4llmStrategy',
    # 转图片策略
    'LayoutToPngStrategy',
    'LayoutToJpgStrategy',
    'LayoutToTifStrategy',
    # 转PDF策略
    'LayoutToPdfStrategy',
    'OfdToPdfStrategy',
    'XpsToPdfStrategy',
    # 操作策略
    'MergePdfsStrategy',
    'SplitPdfStrategy',
]
