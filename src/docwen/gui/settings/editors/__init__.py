"""
编辑器对话框子包

提供各种配置编辑器对话框，继承自 BaseEditorDialog 基类。
"""

from docwen.gui.settings.editors.base import BaseEditorDialog
from docwen.gui.settings.editors.numbering_add import HeadingNumberingEditorDialog
from docwen.gui.settings.editors.numbering_clean import NumberingPatternsEditorDialog

__all__ = [
    "BaseEditorDialog",
    "HeadingNumberingEditorDialog",
    "NumberingPatternsEditorDialog",
]
