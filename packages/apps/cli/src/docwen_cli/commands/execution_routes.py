"""Canonical mapping between public execution commands and runtime actions."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class ExecutionRoute:
    public_path: str
    action: str


STATIC_EXECUTION_ROUTES: dict[str, ExecutionRoute] = {
    "validate": ExecutionRoute("validate", "validate"),
    "number markdown": ExecutionRoute("number markdown", "process_md_numbering"),
    "merge pdf": ExecutionRoute("merge pdf", "merge_pdfs"),
    "merge tables": ExecutionRoute("merge tables", "merge_tables"),
    "merge images": ExecutionRoute("merge images", "merge_images_to_tiff"),
    "split pdf": ExecutionRoute("split pdf", "split_pdf"),
    "batch validate": ExecutionRoute("batch validate", "validate"),
}


def route_for_public_command(
    public_path: str,
) -> ExecutionRoute:
    """Resolve one parser-owned public path to action intent only.

    Named-action targets belong to the canonical Runtime route catalog.  The
    parser never carries a second source/target capability table.
    """

    if public_path in {"convert", "batch convert"}:
        return ExecutionRoute(public_path, "")
    return STATIC_EXECUTION_ROUTES[public_path]


__all__ = [
    "STATIC_EXECUTION_ROUTES",
    "ExecutionRoute",
    "route_for_public_command",
]
