"""Keep spreadsheet drafts outside openpyxl until image markers are consumed."""

from __future__ import annotations

import re
from typing import Any

from docwen_core.text.placeholders import placeholder_namespace, restore_placeholders
from docwen_plugin_markdown.xlsx_text import assign_bounded_xlsx_value, existing_xlsx_cells


class XlsxImageTextPlan:
    """Track cell objects so overwritten or merged-away drafts do not survive."""

    def __init__(self) -> None:
        self._values: dict[Any, Any] = {}
        self._literals: dict[str, str] = {}
        self._template_images: dict[str, str] = {}
        self._authored = ""
        self._image_prefix = ""
        self._template_namespace = ""
        self._template_pattern = re.compile(r"(?!)")

    def value(self, cell: Any) -> Any:
        return self._values.get(cell, cell.value)

    def assign(self, cell: Any, value: Any) -> None:
        # Scope and URL quoting are transport overhead, never cell contents.
        self._values[cell] = value

    def protect_template_literals(self, workbook: Any, *, image_scope: str, markdown: str) -> None:
        self._image_prefix = f"{{{{IMAGE@{image_scope}:"
        cells = [cell for sheet in workbook.worksheets for cell in existing_xlsx_cells(sheet)]
        self._authored = markdown + "".join(str(cell.value) for cell in cells)
        self._template_namespace = image_scope + placeholder_namespace(self._authored, prefix="DOCWENXLSXIMAGE")
        self._template_pattern = re.compile(r"\{\{IMAGE:" + re.escape(self._template_namespace) + r"\d+\}\}")
        for cell in cells:
            if isinstance(cell.value, str) and self._image_prefix in cell.value:
                # A marker already present in the template is not a producer's
                # capability, even if its spelling happens to match this scope.
                self.assign(cell, self.protect_literal_value(cell.value))

    def protect_literal_value(self, value: Any) -> Any:
        if not isinstance(value, str):
            return value
        literals = {
            match.group() for match in self._template_pattern.finditer(value) if match.group() in self._template_images
        }
        if self._image_prefix and self._image_prefix in value:
            literals.add(self._image_prefix)
        for literal in sorted(literals):
            authored = self._authored + value + "".join(self._literals)
            token = placeholder_namespace(authored, prefix="DOCWENXLSXLITERAL") + "IMAGEPREFIX"
            self._literals[token] = literal
            value = value.replace(literal, token)
        return value

    def yaml_parts(self, value: str) -> list[tuple[str, bool]]:
        """Keep original image slots apart from text assembled by YAML fields."""
        parts: list[tuple[str, bool]] = []
        start = 0
        for match in self._template_pattern.finditer(value):
            if match.group() not in self._template_images:
                continue
            parts.extend([(value[start : match.start()], False), (match.group(), True)])
            start = match.end()
        parts.append((value[start:], False))
        return parts

    def template_image(self, marker: str) -> str:
        # Keep the image-shaped grammar during field discovery, but do not
        # expose the trusted capability to authored YAML substitutions.
        token = "{{IMAGE:" + self._template_namespace + str(len(self._template_images)) + "}}"
        self._template_images[token] = marker
        return token

    def image_value(self, cell: Any) -> Any:
        value = self.value(cell)
        if not isinstance(value, str):
            return value
        return self._template_pattern.sub(lambda match: self._template_images.get(match.group(), match.group()), value)

    def final_value(self, value: Any) -> Any:
        return restore_placeholders(value, self._literals) if isinstance(value, str) else value

    def commit(self, cell: Any, value: Any) -> None:
        if cell in self._values or value != cell.value:
            assign_bounded_xlsx_value(cell, value)
        self._values.pop(cell, None)


def image_scope_without_literal_collision(scope: str, markdown: str) -> str:
    """Retain an unpredictable capability while excluding authored marker text."""
    while scope in markdown:
        scope += "_"
    return scope
