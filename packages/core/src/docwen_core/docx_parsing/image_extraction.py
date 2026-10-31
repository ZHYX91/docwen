"""Element-level image extraction from DOCX paragraphs, tables, SDT, and cells."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from docwen_core.docx_parsing.xml_ns import NS_A, NS_R, NS_V, NS_WP


@dataclass(frozen=True)
class ImageInfo:
    alt: str
    path: str
    rel_id: str
    content_type: str
    source_name: str = ""


class ImageExtractionNames:
    """Request-owned names for every DOCX image occurrence in one directory.

    A relationship may recur in distinct paragraphs. Source names may collide
    with each other or with a previously allocated suffix, so neither is a
    unique resource identity. The returned path is reserved for one occurrence.
    """

    def __init__(self, output_dir: str) -> None:
        self.directory = Path(output_dir)
        self.directory.mkdir(parents=True, exist_ok=True)
        self._used = {path.name.casefold() for path in self.directory.iterdir()}
        self._next_suffix: dict[str, int] = {}

    def allocate(self, name: str) -> Path:
        path = Path(name)
        key = name.casefold()
        suffix = self._next_suffix.get(key, 0)
        while True:
            candidate = name if suffix == 0 else f"{path.stem}_{suffix}{path.suffix}"
            result = self.directory / candidate
            if candidate.casefold() not in self._used and not result.exists() and not result.is_symlink():
                self._used.add(candidate.casefold())
                self._next_suffix[key] = suffix + 1
                return result
            suffix += 1


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
    name_allocator: ImageExtractionNames | None = None,
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
    names = name_allocator or ImageExtractionNames(output_dir)
    if names.directory != dest:
        raise ValueError("image name allocator belongs to a different output directory")

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

    # A paragraph or cell can alternate DrawingML and VML occurrences. Walk
    # their real XML order once rather than concatenating two format groups.
    for image_element in element.iter():
        drawingml = image_element.tag == f"{{{NS_A}}}blip"
        if drawingml:
            rel_id = image_element.get(f"{{{NS_R}}}embed") or image_element.get(f"{{{NS_R}}}link")
        elif image_element.tag == f"{{{NS_V}}}imagedata":
            rel_id = image_element.get(f"{{{NS_R}}}id")
        else:
            continue
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
        if drawingml:
            # Walk up to find docPr
            drawing_elem = image_element
            for _ in range(10):  # safety limit
                drawing_elem = drawing_elem.getparent() if hasattr(drawing_elem, "getparent") else None  # type: ignore[assignment]
                if drawing_elem is None:
                    break
                doc_pr = drawing_elem.find(f"{{{NS_WP}}}docPr")
                if doc_pr is not None:
                    description = doc_pr.get("descr")
                    name = description if description is not None else doc_pr.get("name")
                    if name is not None:
                        alt = name
                    source_name = doc_pr.get("title") or ""
                    break

        counter += 1
        fname = _safe_source_filename(source_name, ext) if prefer_source_names else ""
        if not fname:
            fname = f"{name_prefix}_{counter}_{rel_id}{ext}"
        fpath = names.allocate(fname)
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
