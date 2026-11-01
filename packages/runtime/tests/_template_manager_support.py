"""Minimal native template and writable registry fixtures."""

from __future__ import annotations

import zipfile
from pathlib import Path

from docwen_runtime.templates import (
    TemplateManager,
    TemplateRegistry,
    TemplateStateStore,
)


def _write_ooxml_template(path: Path, target: str) -> None:
    main_part = "word/document.xml" if target == "docx" else "xl/workbook.xml"
    content_type = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
        if target == "docx"
        else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    )
    root = "document" if target == "docx" else "workbook"
    namespace = (
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        if target == "docx"
        else "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w") as package:
        package.writestr(
            "[Content_Types].xml",
            (
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                f'<Override PartName="/{main_part}" ContentType="{content_type}"/>'
                "</Types>"
            ),
        )
        package.writestr(
            "_rels/.rels",
            (
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
                f'Target="{main_part}"/>'
                "</Relationships>"
            ),
        )
        package.writestr(main_part, f'<{root} xmlns="{namespace}"/>')


def _manager(tmp_path: Path) -> tuple[TemplateManager, Path, Path]:
    builtin_dir = tmp_path / "bundled" / "templates"
    user_dir = tmp_path / "user" / "templates"
    builtin_dir.mkdir(parents=True)
    user_dir.mkdir(parents=True)
    state = TemplateStateStore(tmp_path / "user" / "template-state.json", user_dir=user_dir)
    registry = TemplateRegistry(
        builtin_dir,
        extra_paths=[user_dir],
        state_store=state,
        managed_user_dir=user_dir,
        builtin_dir=builtin_dir,
    )
    return TemplateManager(registry, state_store=state, user_dir=user_dir), builtin_dir, user_dir
