"""CLI output presenters: text and JSON.

These presenters format ``ConversionResult`` and batch summaries for
terminal output.  They do NOT import plugin or runtime internals.
"""

from docwen_cli.presenters.json_presenter import JsonPresenter
from docwen_cli.presenters.text_presenter import TextPresenter

__all__ = ["JsonPresenter", "TextPresenter"]
