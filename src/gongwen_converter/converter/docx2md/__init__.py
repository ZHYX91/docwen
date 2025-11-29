"""
DOCX转MD子包入口
按需加载转换功能
使用analyze_document_format作为主接口
"""

import logging
import sys

# 配置日志
logger = logging.getLogger(__name__)

__all__ = ['analyze_document_format']

# 模块级缓存
_loaded_functions = {}

def __getattr__(name):
    if name == 'analyze_document_format':
        if name not in _loaded_functions:
            # 安全导入函数
            from .core import analyze_document_format as func
            logger.debug("按需加载DOCX转MD功能")
            _loaded_functions[name] = func
        return _loaded_functions[name]
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")

# 防止递归导入的守卫
if '_loading' not in sys.modules:
    sys.modules['_loading'] = True
    # 可选：预加载一些常用功能