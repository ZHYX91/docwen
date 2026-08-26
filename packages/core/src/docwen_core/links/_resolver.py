"""File path resolution and link-target normalisation.

Covers:
- normalize_link_target: URL-decode and strip anchors / query params
- resolve_file_path: resolve an embedded file path against a source file
"""

from __future__ import annotations

import logging
import os
from pathlib import Path
from urllib.parse import unquote

from PIL import Image

from docwen_core.detection import detect_content_format
from docwen_core.formats import CATEGORY_IMAGE, get_category

logger = logging.getLogger(__name__)


def normalize_link_target(
    link_target: str,
    *,
    preserve_absolute: bool = False,
) -> str:
    """Normalise a link target for consistent Markdown output.

    Normalisation steps (order matters):

    1. Strip raw anchor/query delimiters before decoding so encoded literal
       ``%23`` / ``%3F`` characters remain part of the filename.
    2. Percent-decode (``%20`` → `` ``).
    3. Replace Windows backslashes with forward slashes (``\\`` → ``/``).
    4. By default strip leading slashes for the historical Markdown-output
       href contract.  Set *preserve_absolute* when resolving filesystem
       targets so local absolute and UNC roots survive.
    5. Trim surrounding whitespace.

    Returns the bare, normalised file path portion of the target.
    """
    # 1. Strip structural anchor/query delimiters while still encoded.
    delimiter_positions = [position for delimiter in ("#", "?") if (position := link_target.find(delimiter)) != -1]
    path_end = min(delimiter_positions, default=len(link_target))
    result = link_target[:path_end]

    # 2. Percent-decode exactly once.
    result = unquote(result)

    # 3. Normalise Windows separators
    result = result.replace("\\", "/")

    # 4. Markdown-output consumers historically require relative hrefs, while
    # filesystem resolution must preserve absolute roots.  Protocol-relative
    # URLs are rejected by the shared embed dispatcher before the latter path.
    if not preserve_absolute:
        result = result.lstrip("/")

    # 5. Trim whitespace
    return result.strip()


def resolve_file_path(
    link_target: str,
    source_file_path: str,
    search_dirs: list[str] | None = None,
    *,
    decoded_path: bool = False,
    preserve_absolute: bool = False,
) -> str | None:
    r"""Resolve an embedded file path referenced from *source_file_path*.

    Resolution priority (first match wins):

    1. **Absolute path** — if *link_target* is already absolute.
    2. **Relative path** — relative to the directory of *source_file_path*
       (detected when *link_target* contains ``/`` or ``\\``).
    3. **Same-name folder** — look in ``{source_stem}/`` sibling directory.
    4. **Search directories** — look in each directory in *search_dirs*
       (defaults to ``[\".\", \"assets\", \"images\", \"attachments\"]``).

    Automatic ``.md`` extension: when *link_target* has no recognised
    suffix a ``.md`` variant is tried at each tier, matching Obsidian's
    extension-less wiki-link convention.

    Returns the normalised absolute path as a string, or *None* when the
    file cannot be found.
    """
    import contextlib

    if search_dirs is None:
        search_dirs = [".", "assets", "images", "attachments"]

    logger.debug("Resolving file path: %s", link_target)

    if decoded_path:
        target = link_target.replace("\\", "/").strip()
        if not preserve_absolute and not target.startswith("//"):
            target = target.lstrip("/")
    else:
        target = normalize_link_target(link_target, preserve_absolute=True)
    logger.debug("Normalised target: %s", target)

    source_dir = Path(source_file_path).parent

    def _try_find_file(path: str, *, allowed_root: Path | None = None) -> str | None:
        """Check if *path* exists, or if adding ``.md`` makes it exist."""
        candidate = Path(path)

        def _admit(value: Path) -> str | None:
            with contextlib.suppress(OSError, ValueError):
                if not value.exists():
                    return None
                if allowed_root is not None:
                    value.resolve(strict=True).relative_to(allowed_root.resolve(strict=True))
                return str(value)
            return None

        admitted = _admit(candidate)
        if admitted is not None:
            return admitted
        if not candidate.suffix:
            md_path = candidate.with_suffix(".md")
            admitted = _admit(md_path)
            if admitted is not None:
                logger.debug("Auto-added .md extension: %s -> %s", candidate, md_path)
                return admitted
        return None

    # 1. Absolute path
    if Path(target).is_absolute():
        normalized = os.path.normpath(target)
        result = _try_find_file(normalized)
        if result:
            logger.debug("Resolved as absolute path: %s", result)
            return result
        logger.debug("Absolute path does not exist: %s", normalized)
        return None

    # 2. Relative path (contains directory separator)
    if "/" in target or "\\" in target:
        full_path = os.path.normpath(str(source_dir / target))
        result = _try_find_file(full_path)
        if result:
            logger.debug("Resolved as relative path: %s", result)
            return result
        logger.debug("Relative path does not exist: %s", full_path)
        return None

    # 3. Same-name folder
    source_basename = Path(source_file_path).stem
    same_name_folder = source_dir / source_basename
    if same_name_folder.is_dir():
        search_path = os.path.normpath(str(same_name_folder / target))
        result = _try_find_file(search_path, allowed_root=source_dir)
        if result:
            logger.debug("Found in same-name folder '%s': %s", source_basename, result)
            return result

    # 4. Search directories
    logger.debug("Searching for file: %s (in dirs: %s)", target, search_dirs)
    for search_dir in search_dirs:
        normalized_dir = search_dir.replace("\\", "/").strip()
        search_parts = Path(normalized_dir).parts
        if not normalized_dir or Path(normalized_dir).is_absolute() or any(part in {"..", ""} for part in search_parts):
            logger.warning("Ignoring search directory outside the source-local boundary: %s", search_dir)
            continue
        search_path = os.path.normpath(str(source_dir / normalized_dir / target))
        result = _try_find_file(search_path, allowed_root=source_dir)
        if result:
            logger.debug("Found in search dir '%s': %s", search_dir, result)
            return result

    logger.debug("File not found in any search directory: %s", target)
    return None


def get_file_type(file_path: str) -> str:
    """Determine the embeddable category of *file_path* from its content.

    Returns one of ``"image"``, ``"markdown"``, or ``"unknown"``.

    Embedded resources are an internal conversion boundary, not a top-level
    user admission boundary.  They therefore consume the shared content-only
    detector directly: a filename can help locate a resource, but can never
    grant it an executable image or Markdown route.
    """
    try:
        detected_format = detect_content_format(file_path).format
    except OSError:
        return "unknown"

    category = get_category(detected_format)
    if category == CATEGORY_IMAGE:
        try:
            with Image.open(file_path) as image:
                image.load()
        except Exception:
            # A magic prefix is not sufficient for an embedded resource: the
            # downstream image route must never receive a corrupt payload.
            return "unknown"
        return "image"
    if detected_format in {"markdown", "txt"}:
        return "markdown"
    return "unknown"
