"""Stable invalid-input diagnostics for Word templates."""


class TemplatePackageError(ValueError):
    """A malformed or unsupported template cannot be filled safely."""

    diagnostic_code = "MD2DOCX-TEMPLATE-PACKAGE-INVALID"
    error_type = "invalid_input"


class BodyPlaceholderPlacementError(TemplatePackageError):
    """A body marker exists outside the supported single-paragraph anchor."""

    diagnostic_code = "MD2DOCX-TEMPLATE-BODY-PLACEMENT"
