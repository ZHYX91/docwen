"""Canonical format categories and their member formats."""

from __future__ import annotations

# ── Format categories ────────────────────────────────────────────────

CATEGORY_DOCUMENT = "document"
CATEGORY_SPREADSHEET = "spreadsheet"
CATEGORY_IMAGE = "image"
CATEGORY_LAYOUT = "layout"
CATEGORY_MARKDOWN = "markdown"
CATEGORY_PRESENTATION = "presentation"
CATEGORY_MARKUP = "markup"
CATEGORY_OTHER = "other"

ALL_CATEGORIES: frozenset[str] = frozenset(
    {
        CATEGORY_DOCUMENT,
        CATEGORY_SPREADSHEET,
        CATEGORY_IMAGE,
        CATEGORY_LAYOUT,
        CATEGORY_MARKDOWN,
        CATEGORY_PRESENTATION,
        CATEGORY_MARKUP,
        CATEGORY_OTHER,
    }
)

# ── Format → category mapping ────────────────────────────────────────

FORMAT_CATEGORY: dict[str, str] = {
    # Document family
    "docx": CATEGORY_DOCUMENT,
    "doc": CATEGORY_DOCUMENT,
    "odt": CATEGORY_DOCUMENT,
    "rtf": CATEGORY_DOCUMENT,
    "wps": CATEGORY_DOCUMENT,
    "txt": CATEGORY_DOCUMENT,
    # Spreadsheet family
    "xlsx": CATEGORY_SPREADSHEET,
    "xls": CATEGORY_SPREADSHEET,
    "ods": CATEGORY_SPREADSHEET,
    "et": CATEGORY_SPREADSHEET,
    "csv": CATEGORY_SPREADSHEET,
    "tsv": CATEGORY_SPREADSHEET,
    # Image family
    "png": CATEGORY_IMAGE,
    "jpg": CATEGORY_IMAGE,
    "jpeg": CATEGORY_IMAGE,
    "gif": CATEGORY_IMAGE,
    "bmp": CATEGORY_IMAGE,
    "tif": CATEGORY_IMAGE,
    "tiff": CATEGORY_IMAGE,
    "webp": CATEGORY_IMAGE,
    "heic": CATEGORY_IMAGE,
    "heif": CATEGORY_IMAGE,
    # Layout / fixed-layout family
    "pdf": CATEGORY_LAYOUT,
    "ofd": CATEGORY_LAYOUT,
    "xps": CATEGORY_LAYOUT,
    # Markdown
    "md": CATEGORY_MARKDOWN,
    "markdown": CATEGORY_MARKDOWN,
    # HTML family → markup
    "html": CATEGORY_MARKUP,
    "htm": CATEGORY_MARKUP,
    "mhtml": CATEGORY_MARKUP,
    "mht": CATEGORY_MARKUP,
    # E-book / note export → markup
    "epub": CATEGORY_MARKUP,
    "enex": CATEGORY_MARKUP,
    # Presentation
    "pptx": CATEGORY_PRESENTATION,
    "ppt": CATEGORY_PRESENTATION,
}

# ── Media types ──────────────────────────────────────────────────────

FORMAT_MEDIA_TYPE: dict[str, str] = {
    # Markdown / text
    "md": "text/markdown",
    "markdown": "text/markdown",
    "txt": "text/plain",
    "csv": "text/csv",
    "tsv": "text/tab-separated-values",
    # Document
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "doc": "application/msword",
    "odt": "application/vnd.oasis.opendocument.text",
    "rtf": "application/rtf",
    "wps": "application/vnd.ms-works",
    # Spreadsheet
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "xls": "application/vnd.ms-excel",
    "ods": "application/vnd.oasis.opendocument.spreadsheet",
    # PDF / fixed-layout
    "pdf": "application/pdf",
    "ofd": "application/vnd.ofd",
    "xps": "application/vnd.ms-xpsdocument",
    # Image
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "tif": "image/tiff",
    "tiff": "image/tiff",
    "webp": "image/webp",
    "heic": "image/heic",
    "heif": "image/heif",
    # HTML
    "html": "text/html",
    "htm": "text/html",
    "mhtml": "multipart/related",
    "mht": "multipart/related",
    # E-book / presentation
    "epub": "application/epub+zip",
    "enex": "application/x-evernote",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "ppt": "application/vnd.ms-powerpoint",
}


def get_category(fmt: str) -> str:
    """Return the category for *fmt*, or ``"other"`` if unknown."""
    return FORMAT_CATEGORY.get(fmt.lower(), CATEGORY_OTHER)


def get_media_type(fmt: str) -> str:
    """Return the IANA media type for *fmt*, or a generic octet-stream fallback."""
    return FORMAT_MEDIA_TYPE.get(fmt.lower(), "application/octet-stream")
