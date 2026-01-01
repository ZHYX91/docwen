"""
DOCX转MD子包入口

模块结构:
├── core.py                    # 路由入口
├── shared/                    # 共享处理器
│   ├── table_processor.py     # 表格处理
│   ├── image_processor.py     # 图片处理
│   ├── formula_processor.py   # 公式处理
│   ├── note_processor.py      # 脚注/尾注处理
│   ├── textbox_processor.py   # 文本框处理
│   ├── content_injector.py    # 内容注入
│   ├── list_processor.py      # 列表处理
│   └── break_processor.py     # 边框组处理
├── simple/                    # 通用简化模式
│   └── converter.py           # 简化模式转换器
└── gongwen/                   # 公文文种
    ├── converter.py           # 公文模式转换器
    └── scorer.py              # 公文元素评分器

使用:
    from gongwen_converter.converter.docx2md import convert_docx_to_md
    from gongwen_converter.converter.docx2md import NoteExtractor
"""

import logging
import sys

# 配置日志
logger = logging.getLogger(__name__)

__all__ = [
    # 核心转换函数
    'convert_docx_to_md',
    # 脚注/尾注处理（向后兼容）
    'NoteExtractor',
    'extract_footnotes_from_docx',
    'extract_endnotes_from_docx',
    'build_footnote_definitions',
    'build_endnote_definitions',
]

# 导出脚注/尾注处理类（从新位置）
from .shared.note_processor import (
    NoteExtractor,
    extract_footnotes_from_docx,
    extract_endnotes_from_docx,
    build_footnote_definitions,
    build_endnote_definitions,
)

# 模块级缓存
_loaded_functions = {}

def __getattr__(name):
    """
    延迟加载模块函数
    
    支持的名称:
    - convert_docx_to_md: 主转换函数（路由到简化或公文模式）
    - analyze_document_format: 文档格式分析（向后兼容别名）
    """
    if name in ('convert_docx_to_md', 'analyze_document_format'):
        if name not in _loaded_functions:
            from .core import convert_docx_to_md as func
            logger.debug("按需加载DOCX转MD功能")
            _loaded_functions['convert_docx_to_md'] = func
            _loaded_functions['analyze_document_format'] = func  # 别名
        return _loaded_functions.get(name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")

# 防止递归导入的守卫
if '_loading' not in sys.modules:
    sys.modules['_loading'] = True
