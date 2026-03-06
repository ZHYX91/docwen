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
    from docwen.converter.docx2md import convert_docx_to_md
    from docwen.converter.docx2md import NoteExtractor
"""

import logging

from .shared.note_processor import (
    NoteExtractor,
    build_endnote_definitions,
    build_footnote_definitions,
    extract_endnotes_from_docx,
    extract_footnotes_from_docx,
)

# 配置日志
logger = logging.getLogger(__name__)

__all__ = [
    # 脚注/尾注处理（向后兼容）
    "NoteExtractor",
    "build_endnote_definitions",
    "build_footnote_definitions",
    # 核心转换函数
    "convert_docx_to_md",
    "extract_endnotes_from_docx",
    "extract_footnotes_from_docx",
]

# 模块级缓存
_loaded_functions = {}


def convert_docx_to_md(*args, **kwargs):
    from .core import convert_docx_to_md as func

    return func(*args, **kwargs)


def __getattr__(name):
    """
    延迟加载模块函数

    支持的名称:
    - convert_docx_to_md: 主转换函数（路由到简化或公文模式）
    """
    if name == "convert_docx_to_md":
        if name not in _loaded_functions:
            logger.debug("按需加载DOCX转MD功能")
            _loaded_functions["convert_docx_to_md"] = convert_docx_to_md
        return _loaded_functions.get(name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
