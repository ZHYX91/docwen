"""Element-level image extraction from DOCX paragraphs, tables, SDT, and cells."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from docwen_core.docx_parsing.xml_ns import NS_R, NS_WP, NSMAP_WPV


@dataclass(frozen=True)
class ImageInfo:
    alt: str
    path: str
    rel_id: str
    content_type: str
    source_name: str = ""


def _content_type_to_ext(content_type: str) -> str | None:
    """Map MIME content type to file extension; return None for EMF/WMF."""
    mapping = {
        "image/png": ".png",
        "image/jpeg": ".jpg",
        "image/gif": ".gif",
        "image/bmp": ".bmp",
        "image/tiff": ".tiff",
        "image/svg+xml": ".svg",
    }
    return mapping.get(content_type)


def extract_images_from_element(
    element,
    related_parts,
    output_dir: str,
    *,
    name_prefix: str = "docx-image",
    prefer_source_names: bool = False,
) -> list[ImageInfo]:
    """Extract images from any XML element (paragraph, table, SDT, textbox, cell).

    Searches DrawingML ``w:drawing//a:blip`` (with ``r:embed`` or ``r:link``)
    and VML ``w:pict//v:imagedata`` (with ``r:id``).  Filters EMF/WMF via
    ``_content_type_to_ext()``.  Writes image files to *output_dir* and returns
    ``ImageInfo`` records.

    Args:
        element: An lxml element (paragraph, table, etc.).
        related_parts: A mapping from relationship id to ``ImagePart`` (e.g.
                       ``doc.part.related_parts`` or a header/footer part).
        output_dir: Directory path for extracted image files.
        name_prefix: Filename prefix (default ``"docx-image"``).

    Returns:
        List of ``ImageInfo`` objects for successfully extracted images.
    """
    dest = Path(output_dir)
    dest.mkdir(parents=True, exist_ok=True)

    infos: list[ImageInfo] = []
    created_paths: list[Path] = []
    counter = 0

    def write_image(path: Path, content: bytes) -> None:
        """Write one extraction atomically from the caller's perspective."""

        created_paths.append(path)
        try:
            path.write_bytes(content)
        except BaseException:
            for created_path in created_paths:
                created_path.unlink(missing_ok=True)
            raise

    # ── DrawingML (w:drawing) ────────────────────────────────────────
    for blip in element.findall(".//a:blip", NSMAP_WPV):
        rel_id = blip.get(f"{{{NS_R}}}embed") or blip.get(f"{{{NS_R}}}link")
        if not rel_id:
            continue

        try:
            image_part = related_parts[rel_id]
            image_bytes = image_part.blob
        except (KeyError, AttributeError):
            continue

        content_type = getattr(image_part, "content_type", "")
        ext = _content_type_to_ext(content_type)
        if ext is None:
            continue  # skip EMF/WMF

        # Alt text from drawing properties
        alt = "image"
        source_name = ""
        # Walk up to find docPr
        drawing_elem = blip
        for _ in range(10):  # safety limit
            drawing_elem = drawing_elem.getparent() if hasattr(drawing_elem, "getparent") else None  # type: ignore[assignment]
            if drawing_elem is None:
                break
            doc_pr = drawing_elem.find(f"{{{NS_WP}}}docPr")
            if doc_pr is not None:
                name = doc_pr.get("descr") or doc_pr.get("name") or ""
                if name:
                    alt = name
                source_name = doc_pr.get("title") or ""
                break

        counter += 1
        fname = _safe_source_filename(source_name, ext) if prefer_source_names else ""
        if not fname:
            fname = f"{name_prefix}_{counter}_{rel_id}{ext}"
        fpath = dest / fname
        if fpath.exists():
            fpath = dest / f"{fpath.stem}_{counter}{fpath.suffix}"
        write_image(fpath, image_bytes)
        infos.append(
            ImageInfo(
                alt=alt,
                path=str(fpath),
                rel_id=rel_id,
                content_type=content_type,
                source_name=source_name,
            )
        )

    # ── VML (w:pict) ─────────────────────────────────────────────────
    for imagedata in element.findall(".//v:imagedata", NSMAP_WPV):
        rel_id = imagedata.get(f"{{{NS_R}}}id")
        if not rel_id:
            continue

        try:
            image_part = related_parts[rel_id]
            image_bytes = image_part.blob
        except (KeyError, AttributeError):
            continue

        content_type = getattr(image_part, "content_type", "")
        ext = _content_type_to_ext(content_type)
        if ext is None:
            continue

        alt = "image"
        counter += 1
        fname = f"{name_prefix}_{counter}_{rel_id}{ext}"
        fpath = dest / fname
        write_image(fpath, image_bytes)
        infos.append(ImageInfo(alt=alt, path=str(fpath), rel_id=rel_id, content_type=content_type))

    return infos


def _safe_source_filename(source_name: str, extension: str) -> str:
    """Return one basename-only resource name suitable for staging."""

    candidate = Path(source_name.replace("\\", "/")).name.strip()
    if not candidate or candidate in {".", ".."}:
        return ""
    if any(character in candidate for character in '<>:"/\\|?*'):
        return ""
    path = Path(candidate)
    if path.suffix.lower() != extension.lower():
        candidate = f"{path.stem or 'image'}{extension}"
    return candidate
