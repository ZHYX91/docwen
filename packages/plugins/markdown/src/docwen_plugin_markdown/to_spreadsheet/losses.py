"""Occurrence-owned spreadsheet loss diagnostics and non-destructive comments."""

from __future__ import annotations

from typing import Any

from openpyxl.comments import Comment

from docwen_core.conversion_loss import ConversionLoss, LossCollector, loss_text
from docwen_core.models.result import ConversionDiagnostic
from docwen_plugin_markdown.mistune_extensions import create_extended_markdown


class SpreadsheetLosses:
    def __init__(
        self, *, comments: bool, messages: dict[str, str] | None = None, include_source_content: bool = True
    ) -> None:
        self.comments = comments
        self.messages = messages or {}
        self.include_source_content = include_source_content
        self.collector = LossCollector()
        self.annotated_count = 0
        self.annotation_failures: list[ConversionDiagnostic] = []
        self.parser = create_extended_markdown()
        self.written_cells: dict[tuple[str, str], tuple[Any, Any]] = {}

    @property
    def diagnostics(self):
        details = (
            *(
                loss.to_diagnostic(self.messages, include_source_content=self.include_source_content)
                for loss in self.collector.records
            ),
            *self.annotation_failures,
        )
        summary = self.collector.summary(self.messages, annotated_count=self.annotated_count)
        return (*details, summary) if summary else details

    def writer(self, write_cell, project):
        def write(cell, value):
            write_cell(cell, value)
            self.written_cells[(cell.parent.title, cell.coordinate)] = (cell, project(value))

        return write

    def record_written_formulas(self) -> None:
        for cell, value in self.written_cells.values():
            self.formulas(value, f"{cell.parent.title}!{cell.coordinate}", cell if self.comments else None)

    def record(self, cell: Any, loss: ConversionLoss) -> None:
        if not self.collector.add(loss) or not self.comments:
            return
        try:
            previous = cell.comment
            text = loss.message(self.messages)
            if previous is not None:
                # openpyxl has one note per cell. Keep the original author's
                # note and append the conversion explanation, never replace it.
                text = previous.text + "\n\nDocWen: " + text
                comment = Comment(text, previous.author)
                comment.width, comment.height = previous.width, previous.height
            else:
                comment = Comment(text, "DocWen")
            cell.comment = comment
            self.annotated_count += 1
        except Exception:
            self.annotation_failures.append(
                ConversionLoss(
                    code="CONVERSION-LOSS-COMMENT-FAILED",
                    reason=loss_text(self.messages, "comment_failed_reason", "The cell comment could not be written."),
                    change=loss_text(
                        self.messages,
                        "comment_failed_change",
                        "The conversion loss remains available in conversion details.",
                    ),
                    output_location=loss.output_location,
                ).to_diagnostic(self.messages, include_source_content=self.include_source_content)
            )

    def protected(self, cell: Any, original: str) -> None:
        self.record(
            cell,
            ConversionLoss(
                code="CONVERSION-LOSS-PROTECTED-CELL",
                reason=loss_text(self.messages, "protected_cell_reason", "The template cell is protected."),
                change=loss_text(
                    self.messages,
                    "protected_cell_change",
                    "The incoming value was not written; the existing cell was retained.",
                ),
                source_location="Markdown table value",
                output_location=f"{cell.parent.title}!{cell.coordinate}",
                original=original,
            ),
        )

    def formulas(self, value: Any, location: str, cell: Any = None) -> None:
        if not isinstance(value, str) or "$" not in value:
            return
        tokens = self.parser(value)
        if not isinstance(tokens, list):
            return

        def visit(nodes):
            for node in nodes:
                if node.get("type") in {"inline_math", "block_math"}:
                    loss = ConversionLoss(
                        code="CONVERSION-LOSS-SPREADSHEET-MATH",
                        reason=loss_text(
                            self.messages,
                            "math_as_text_reason",
                            "This spreadsheet output does not render mathematical equations.",
                        ),
                        change=loss_text(
                            self.messages, "math_as_text_change", "The equation is retained as readable LaTeX text."
                        ),
                        source_location=location,
                        output_location=location,
                        original=str(node.get("raw", "")),
                    )
                    if cell is None:
                        self.collector.add(loss)
                    else:
                        self.record(cell, loss)
                visit(node.get("children", []))

        visit(tokens)

    def merge_to_csv(self, location: str) -> None:
        self.collector.add(
            ConversionLoss(
                code="CONVERSION-LOSS-CSV-MERGE",
                reason=loss_text(self.messages, "csv_merge_reason", "CSV cannot represent merged cells."),
                change=loss_text(
                    self.messages,
                    "csv_merge_change",
                    "The existing row and column values are retained without merging.",
                ),
                source_location=location,
                output_location=location,
            )
        )

    def merge_failed(self, cell: Any, original: str) -> None:
        self.record(
            cell,
            ConversionLoss(
                code="CONVERSION-LOSS-SPREADSHEET-MERGE",
                reason=loss_text(
                    self.messages, "merge_failed_reason", "The merged-cell structure could not be preserved."
                ),
                change=loss_text(
                    self.messages, "merge_failed_change", "Cell text and existing template merges were retained."
                ),
                output_location=f"{cell.parent.title}!{cell.coordinate}",
                original=original,
            ),
        )
