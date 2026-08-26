"""Core error types and exception hierarchy.

All domain-level errors raised by docwen_core and its consumers
should derive from DocWenError.
"""


class DocWenError(Exception):
    """Base error for all DocWen exceptions."""


class ConversionError(DocWenError):
    """Raised when a conversion fails."""


class ConfigurationError(DocWenError):
    """Raised when configuration is invalid or missing."""


class ValidationError(DocWenError):
    """Raised when input validation fails."""


class CancellationRequested(DocWenError):
    """Raised when a task is cancelled."""
