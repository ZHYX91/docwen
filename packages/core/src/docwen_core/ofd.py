"""Runtime monkey-patches for the easyofd library.

Upstream issues (unfixed as of 2026-04-27):

1. **FileRead.__init__** hardcodes ``os.getcwd()`` as the scratch directory.
   When ``BadZipFile`` or other exceptions occur the cleanup path is
   unreachable, leaking temporary OFD/ZIP artefacts into the CWD.

2. **DrawPDF.draw_annotation** calls ``.get("type")`` on a possibly-None
   ``AnnoType``, crashing with ``AttributeError``.  A single malformed
   annotation kills every annotation on the page — there is no per-item
   ``try/except``.

3. **ContentFileParser.fetch_cell_info** assumes every text clip path has a
   numeric ``@Boundary``.  Real producer output can describe the clip with
   ``ofd:AbbreviatedData`` only; easyofd then evaluates ``float("")`` even
   though its parsed ``clips_pos`` value is not consumed by the renderer.

4. **DrawPDF.OP** uses a 200-DPI pixel conversion for ReportLab coordinates.
   ReportLab uses PDF points, so A4 ``210 x 297 mm`` becomes an oversized
   ``1653.54 x 2338.58 pt`` page instead of ``595.28 x 841.89 pt``.

Patches applied by :func:`apply_easyofd_patches`:

* Ignore only unbounded text-clip paths while parsing, avoiding the upstream
  ``float("")`` crash without changing bounded-clip parsing.
* Use the PDF-standard ``72 / 25.4`` points-per-millimetre scale when drawing.
* Give every ``FileRead`` call one owned temporary directory and bypass the
  upstream ``str.split('.')`` unpack-path derivation.  The upstream code can
  otherwise truncate a dotted parent path and recursively delete an ancestor.
* Shadow upstream module-level ``print`` lookups on the OFD-to-PDF path so
  conversion cannot corrupt a framed-stdio transport.
* Wrap ``draw_annotation`` with per-item error isolation: Watermark/Stamp
  annotations are extracted safely; TextObject and friends are delegated
  to the original method under an outer ``try/except``.
"""

from __future__ import annotations

import base64
import importlib
import io
import logging
import stat
import threading
import warnings
import zipfile
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import Any, cast

logger = logging.getLogger(__name__)

_patches_applied: bool = False
_patch_lock = threading.RLock()


def apply_easyofd_patches() -> None:
    """Apply runtime monkey-patches to easyofd (idempotent).

    Safe to call multiple times — a module-level guard prevents
    double-patching.
    """
    global _patches_applied
    with _patch_lock:
        with easyofd_import_boundary():
            if _patches_applied:
                return
            results = {
                "stdout": _patch_easyofd_stdout(),
                "draw_annotation": _patch_draw_annotation(),
                "page_scale": _patch_draw_pdf_page_scale(),
                "content_clip_boundary": _patch_content_clip_boundary(),
                "fileread": _patch_fileread(),
            }
        missing = sorted(name for name, applied in results.items() if not applied)
        if missing:
            raise RuntimeError(f"Required easyofd compatibility patches are unavailable: {', '.join(missing)}")
        _patches_applied = True


def silence_easyofd_import_logging() -> None:
    """Disable EasyOFD logging before importing EasyOFD itself.

    EasyOFD registers fonts at import time and reports missing host fonts on
    stderr. Machine Protocol requires stderr to remain clean, so conversion
    entry points call this boundary before their first ``easyofd`` import.
    """

    try:
        from loguru import logger as loguru_logger
    except ImportError:
        return
    loguru_logger.disable("easyofd")


@contextmanager
def easyofd_import_boundary() -> Iterator[None]:
    """Keep EasyOFD's import-time diagnostics out of framed transports."""

    silence_easyofd_import_logging()
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", SyntaxWarning)
        yield


_EASYOFD_STDOUT_MODULES = (
    "easyofd.draw.draw_pdf",
    "easyofd.draw.font_tools",
    "easyofd.parser_ofd.file_deal",
    "easyofd.parser_ofd.file_parser",
    "easyofd.parser_ofd.file_parser_base",
    "easyofd.parser_ofd.file_publicres_parser",
)


def _discard_easyofd_print(*_args: object, **_kwargs: object) -> None:
    """Discard upstream debugging prints without redirecting process stdout."""


def _patch_easyofd_stdout() -> bool:
    """Shadow OFD-to-PDF debug prints without changing process-global stdout."""

    try:
        modules = tuple(importlib.import_module(name) for name in _EASYOFD_STDOUT_MODULES)
    except Exception as exc:
        logger.warning("Cannot import easyofd stdout boundary - skipping patch: %s", exc, exc_info=True)
        return False
    for module in modules:
        module.__dict__["print"] = _discard_easyofd_print
    return all(getattr(module, "print", None) is _discard_easyofd_print for module in modules)


# ═══════════════════════════════════════════════════════════════════════════
# draw_annotation — per-item error isolation
# ═══════════════════════════════════════════════════════════════════════════


def _patch_draw_annotation() -> bool:
    try:
        from easyofd.draw.draw_pdf import DrawPDF
    except Exception as exc:
        logger.warning(
            "Cannot import easyofd DrawPDF — skipping annotation patch: %s",
            exc,
            exc_info=True,
        )
        return False

    current = getattr(DrawPDF, "draw_annotation", None)
    if current is None:
        return False
    if getattr(current, "_docwen_patched", False):
        return True

    _original_draw_annotation = DrawPDF.draw_annotation

    def _patched(self, canvas, annota_info, images, page_size):
        if not annota_info:
            return

        img_list: list[dict] = []
        native_annos: dict = {}
        skipped_count = 0
        skipped_reasons: dict[str, int] = {}
        warned_reasons: set[str] = set()

        for key, annotation in annota_info.items():
            try:
                if not annotation:
                    continue

                anno_type_obj = annotation.get("AnnoType")
                if not anno_type_obj:
                    continue

                anno_type = anno_type_obj.get("type")

                if anno_type in ("Watermark", "Stamp"):
                    img_obj = annotation.get("ImageObject")
                    if not img_obj:
                        continue

                    boundary_str = img_obj.get("Boundary") or ""
                    pos_str = boundary_str.split(" ") if boundary_str else []
                    pos = [float(i) for i in pos_str] if pos_str else []

                    appearance = annotation.get("Appearance") or {}
                    wrap_boundary_str = appearance.get("Boundary") or ""
                    wrap_pos_str = wrap_boundary_str.split(" ") if wrap_boundary_str else []
                    wrap_pos = [float(i) for i in wrap_pos_str] if wrap_pos_str else []

                    ctm_str = img_obj.get("CTM") or ""
                    ctm_split = ctm_str.split(" ") if ctm_str else []
                    ctm = [float(i) for i in ctm_split] if ctm_split else []

                    img_list.append(
                        {
                            "wrap_pos": wrap_pos,
                            "pos": pos,
                            "CTM": ctm,
                            "ResourceID": img_obj.get("ResourceID", ""),
                        }
                    )
                else:
                    native_annos[key] = annotation

            except (KeyError, TypeError, ValueError) as exc:
                skipped_count += 1
                key_name = type(exc).__name__
                skipped_reasons[key_name] = skipped_reasons.get(key_name, 0) + 1
                continue
            except Exception as exc:
                skipped_count += 1
                key_name = type(exc).__name__
                skipped_reasons[key_name] = skipped_reasons.get(key_name, 0) + 1
                if key_name not in warned_reasons:
                    warned_reasons.add(key_name)
                    logger.warning(
                        "easyofd annotation structure mismatch / version incompatibility — skipping annotation: %s",
                        exc,
                        exc_info=True,
                    )
                continue

        if img_list and hasattr(self, "draw_img"):
            if skipped_count:
                logger.warning(
                    "easyofd annotation patch skipped %s annotation(s): %s",
                    skipped_count,
                    ", ".join(f"{k}={v}" for k, v in sorted(skipped_reasons.items())),
                )
            self.draw_img(canvas, img_list, images, page_size)

        if native_annos:
            try:
                _original_draw_annotation(self, canvas, native_annos, images, page_size)
            except Exception as e:
                logger.warning(
                    "easyofd original draw_annotation failed (TextObject, …): %s",
                    e,
                    exc_info=True,
                )

    _patched._docwen_patched = True  # type: ignore[attr-defined]
    DrawPDF.draw_annotation = _patched
    return bool(getattr(DrawPDF.draw_annotation, "_docwen_patched", False))


def _patch_content_clip_boundary() -> bool:
    """Tolerate clip paths represented solely by abbreviated path data.

    easyofd only stores ``clips_pos`` and never consumes it while rendering.
    When a producer omits ``@Boundary`` from ``ofd:Path``, remove the clip
    branch from a shallow row copy before delegating to the upstream parser.
    The original row and every bounded clip remain untouched.
    """
    try:
        from easyofd.parser_ofd.file_content_parser import ContentFileParser
    except Exception as exc:
        logger.warning(
            "Cannot import easyofd ContentFileParser - skipping clip-boundary patch: %s",
            exc,
            exc_info=True,
        )
        return False

    current = getattr(ContentFileParser, "fetch_cell_info", None)
    if current is None:
        return False
    if getattr(current, "_docwen_patched", False):
        return True

    original_fetch_cell_info = current

    def _patched(self, row, TextObject):
        path = row.get("ofd:Clips", {}).get("ofd:Clip", {}).get("ofd:Area", {}).get("ofd:Path", {})
        if isinstance(path, dict) and path and not str(path.get("@Boundary") or "").strip():
            row = dict(row)
            row.pop("ofd:Clips", None)
        return original_fetch_cell_info(self, row, TextObject)

    _patched._docwen_patched = True  # type: ignore[attr-defined]
    ContentFileParser.fetch_cell_info = _patched
    return bool(getattr(ContentFileParser.fetch_cell_info, "_docwen_patched", False))


def _patch_draw_pdf_page_scale() -> bool:
    """Use PDF points rather than 200-DPI pixels for OFD millimetres."""
    try:
        from easyofd.draw.draw_pdf import DrawPDF
    except Exception as exc:
        logger.warning(
            "Cannot import easyofd DrawPDF - skipping page-scale patch: %s",
            exc,
            exc_info=True,
        )
        return False

    current = getattr(DrawPDF, "__init__", None)
    if current is None:
        return False
    if getattr(current, "_docwen_page_scale_patched", False):
        return True

    original_init = current

    def _patched_init(self, *args, **kwargs):
        original_init(self, *args, **kwargs)
        self.OP = 72 / 25.4

    _patched_init._docwen_page_scale_patched = True  # type: ignore[attr-defined]
    DrawPDF.__init__ = _patched_init
    return bool(getattr(DrawPDF.__init__, "_docwen_page_scale_patched", False))


# ═══════════════════════════════════════════════════════════════════════════
# FileRead lifecycle — use one exact, owned scratch directory
# ═══════════════════════════════════════════════════════════════════════════


def _patch_fileread() -> bool:
    try:
        from easyofd.parser_ofd import file_deal as file_deal_module
    except Exception as exc:
        logger.warning(
            "Cannot import easyofd FileRead — skipping path patch: %s",
            exc,
            exc_info=True,
        )
        return False

    FileRead = file_deal_module.FileRead

    members = tuple(getattr(FileRead, name, None) for name in ("__init__", "unzip_file", "buld_file_tree", "__call__"))
    if any(member is None for member in members):
        return False
    patched = tuple(bool(getattr(member, "_docwen_fileread_patched", False)) for member in members)
    if all(patched):
        return True
    if any(patched):
        logger.error("Refusing to layer the EasyOFD FileRead patch over a partial prior patch")
        return False
    original_init = cast(Any, members[0])

    def _patched_init(self, ofdb64: str):
        original_init(self, ofdb64)
        import tempfile

        owner_root = Path(tempfile.mkdtemp(prefix="docwen-ofd-")).resolve(strict=True)
        self._docwen_scratch_root = owner_root
        owner_stats = owner_root.lstat()
        self._docwen_scratch_signature = (
            owner_stats.st_dev,
            owner_stats.st_ino,
            owner_stats.st_mode,
            int(getattr(owner_stats, "st_file_attributes", 0)),
        )
        self.zip_path = str(owner_root / "source.ofd")
        self.unzip_path = str(owner_root / "unpacked")

    def _verified_owner(self) -> Path:
        owner_root = Path(self._docwen_scratch_root)
        owner_stats = owner_root.lstat()
        current_signature = (
            owner_stats.st_dev,
            owner_stats.st_ino,
            owner_stats.st_mode,
            int(getattr(owner_stats, "st_file_attributes", 0)),
        )
        reparse_flag = int(getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0))
        if (
            current_signature != self._docwen_scratch_signature
            or not stat.S_ISDIR(owner_stats.st_mode)
            or stat.S_ISLNK(owner_stats.st_mode)
            or (reparse_flag and current_signature[3] & reparse_flag)
        ):
            raise RuntimeError("easyofd scratch owner identity changed")
        return owner_root

    def _cleanup_owner(self) -> None:
        owner_root = _verified_owner(self)
        try:
            owner_root.rmdir()
        except OSError as error:
            raise RuntimeError("easyofd scratch owner is not an empty owned directory") from error

    def _patched_unzip_file(self) -> None:
        owner_root = _verified_owner(self)
        zip_path = Path(self.zip_path)
        unpack_path = Path(self.unzip_path)
        if zip_path.parent != owner_root or unpack_path.parent != owner_root:
            raise RuntimeError("easyofd scratch path escaped its owned root")
        if self.save_xml:
            raise ValueError("easyofd XML export is disabled at the DocWen conversion boundary")

        members: list[tuple[tuple[str, ...], bytes]] = []
        with zipfile.ZipFile(io.BytesIO(self.ofdbyte), "r") as archive:
            for member in archive.infolist():
                member_path = Path(member.filename.replace("\\", "/"))
                if (
                    member_path.is_absolute()
                    or not member_path.parts
                    or any(part in {"", ".", ".."} for part in member_path.parts)
                    or member_path.drive
                ):
                    raise ValueError("unsafe OFD archive member path")
                destination = unpack_path.joinpath(*member_path.parts).resolve(strict=False)
                if destination != unpack_path and unpack_path not in destination.parents:
                    raise ValueError("unsafe OFD archive member path")
                if not member.is_dir():
                    members.append((member_path.parts, archive.read(member)))
        self._docwen_archive_members = members

    def _patched_build_file_tree(self) -> None:
        owner_root = _verified_owner(self)
        unpack_path = Path(self.unzip_path)
        zip_path = Path(self.zip_path)
        if unpack_path.parent != owner_root or zip_path.parent != owner_root:
            raise RuntimeError("easyofd cleanup target escaped its owned root")

        self.file_tree["root"] = str(unpack_path)
        self.file_tree["pdf_name"] = self.pdf_name
        for member_parts, payload in self._docwen_archive_members:
            absolute_path = unpack_path.joinpath(*member_parts)
            file_name = member_parts[-1]
            self.file_tree[str(absolute_path)] = (
                base64.b64encode(payload).decode("utf-8")
                if "xml" not in file_name
                else file_deal_module.xmltodict.parse(payload.decode("utf-8"))
            )
        root_document = unpack_path / "OFD.xml"
        self.file_tree["root_doc"] = str(root_document) if str(root_document) in self.file_tree else ""
        self._docwen_archive_members = []

    def _patched_call(self, *args: object, **kwargs: object):
        self.save_xml = kwargs.get("save_xml", False)
        self.xml_name = kwargs.get("xml_name")
        try:
            self.unzip_file()
            self.buld_file_tree()
        except BaseException as original_error:
            try:
                _cleanup_owner(self)
            except BaseException as cleanup_error:
                original_error.add_note(f"failed to clean owned EasyOFD scratch: {cleanup_error}")
            raise
        _cleanup_owner(self)
        return self.file_tree

    _patched_init._docwen_fileread_patched = True  # type: ignore[attr-defined]
    _patched_unzip_file._docwen_fileread_patched = True  # type: ignore[attr-defined]
    _patched_build_file_tree._docwen_fileread_patched = True  # type: ignore[attr-defined]
    _patched_call._docwen_fileread_patched = True  # type: ignore[attr-defined]
    FileRead.__init__ = _patched_init
    FileRead.unzip_file = _patched_unzip_file
    FileRead.buld_file_tree = _patched_build_file_tree
    FileRead.__call__ = _patched_call
    return all(
        bool(getattr(member, "_docwen_fileread_patched", False))
        for member in (FileRead.__init__, FileRead.unzip_file, FileRead.buld_file_tree, FileRead.__call__)
    )
