"""SmartConverter document-format interconversion (docx/doc/odt/rtf/wps).

Routes are implemented via COM/LibreOffice bridge to external office software
(WPS / Microsoft Office / LibreOffice). Pure Python cannot perform these
format conversions directly.
"""

from __future__ import annotations

from docwen_plugin_document.to_document.converter import SmartDocConverter

__all__ = ["SmartDocConverter"]
