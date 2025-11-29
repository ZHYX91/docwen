"""
模板包入口模块
"""

from .loader import (
    TemplateLoader,
    get_template_dir,
    get_template_path,
    load_docx_template,
    load_xlsx_template,
    get_available_templates
)
from .validator import (
    TemplateValidator,
    validate_template
)

__all__ = [
    'TemplateLoader',
    'TemplateValidator',
    'get_template_dir',
    'get_template_path',
    'load_docx_template',
    'load_xlsx_template',
    'get_available_templates',
    'validate_template'
]

# 包文档
__doc__ = """
公文转换器模板处理包
提供模板加载、验证和管理功能
"""
