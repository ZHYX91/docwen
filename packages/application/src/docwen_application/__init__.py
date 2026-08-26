"""DocWen Application — the admitted application controller and runtime ports.

Depends on: docwen_core
Must NOT depend on: docwen_runtime, docwen_gui, docwen_cli, docwen_plugin_*
"""

from docwen_application.controller import ApplicationController, ControllerError
from docwen_application.conversion_service import ConversionService, ConversionServiceError
from docwen_core.version import __version__ as __version__

__all__ = [
    "ApplicationController",
    "ControllerError",
    "ConversionService",
    "ConversionServiceError",
]
