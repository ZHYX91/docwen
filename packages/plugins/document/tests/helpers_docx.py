from __future__ import annotations

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from docx.document import Document


def save_docx(doc: Document, tmp_path: Path, name: str = "sample.docx") -> Path:
    path = tmp_path / name
    doc.save(str(path))
    return path


def patch_docx_xml(
    src: Path,
    dst: Path,
    replacements: dict[str, str],
    extra_files: dict[str, bytes] | None = None,
) -> Path:
    extra_files = extra_files or {}
    with ZipFile(src, "r") as zin, ZipFile(dst, "w", ZIP_DEFLATED) as zout:
        written = set()
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in replacements:
                text = data.decode("utf-8")
                text = text.replace("</w:body>", replacements[item.filename] + "</w:body>")
                data = text.encode("utf-8")
            zout.writestr(item, data)
            written.add(item.filename)
        for filename, data in extra_files.items():
            if filename not in written:
                zout.writestr(filename, data)
    return dst
