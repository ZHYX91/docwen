"""GUI widget components for the DocWen main window.

Widgets do NOT call runtime/plugins directly — they go through ViewModels
which delegate to ApplicationController.
"""

from .action_area import ActionArea
from .batch_list import BatchList
from .conversion_panel import ConversionPanel
from .info_area import InfoArea
from .input_area import InputArea
from .template_selector import TemplateItemDetails, TemplateSelectionFeedback, TemplateSelector
from .template_selector_tabbed import TabbedTemplateSelector

__all__ = [
    "ActionArea",
    "BatchList",
    "ConversionPanel",
    "InfoArea",
    "InputArea",
    "TabbedTemplateSelector",
    "TemplateItemDetails",
    "TemplateSelectionFeedback",
    "TemplateSelector",
]
