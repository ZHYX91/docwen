"""Project OFD pages onto easyofd's graphics and ReportLab's embedded text."""

from __future__ import annotations

import base64
import logging
import math
import posixpath
import re
import struct
from collections import defaultdict
from collections.abc import Callable, Iterator
from io import BytesIO
from itertools import chain, repeat
from typing import Any, cast
from uuid import uuid4
from zipfile import BadZipFile, ZipFile

import pymupdf as fitz
import xmltodict
from easyofd.draw.draw_pdf import DrawPDF
from easyofd.draw.find_seal_img import SealExtract
from easyofd.parser_ofd.file_content_parser import ContentFileParser
from easyofd.parser_ofd.file_docres_parser import DocumentResFileParser
from easyofd.parser_ofd.file_publicres_parser import PublicResFileParser
from fontTools.pens.basePen import BasePen
from fontTools.svgLib.path import parse_path
from PIL import Image
from pyasn1.codec.der.decoder import decode
from pyasn1.error import PyAsn1Error
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFError, TTFont
from reportlab.pdfgen.canvas import Canvas
from reportlab.pdfgen.pathobject import PDFPathObject

from docwen_core._ofd_glyphs import Outline, OutlineFont

logger = logging.getLogger(__name__)
_NAMESPACE = {
    "http://www.ofdspec.org/2016": "ofd",
    "http://www.ofdspec.org": "ofd",
    # Older easyofd-generated documents use the writer author's URI.
    "http://blog.yuanhaiying.cn": "ofd",
}
_NATIVE_IMAGES = {"jpg", "jpeg", "png"}


def _items(value: Any) -> list[Any]:
    if value is None:
        return []
    return value if isinstance(value, list) else [value]


def _value(value: Any) -> str:
    return (value.get("#text", "") if isinstance(value, dict) else value or "").strip()


def _path(ref: str, base: str = "") -> str:
    """Resolve an OFD ST_Loc against a directory inside its ZIP package."""
    if not ref:
        raise ValueError("Missing OFD member reference")
    ref = ref.replace("\\", "/")
    path = posixpath.normpath(ref.lstrip("/") if ref.startswith("/") else posixpath.join(base, ref))
    if path == ".." or path.startswith("../"):
        raise ValueError(f"OFD member reference escapes the package: {ref}")
    return path


def _read(archive: ZipFile, path: str) -> bytes:
    try:
        return archive.read(path)
    except KeyError as exc:
        raise ValueError(f"OFD member not found: {path}") from exc


def _xml(archive: ZipFile, path: str, root: str) -> dict:
    # xmltodict's final mapping groups siblings by tag. Its postprocessor
    # still visits each graphic object in XML order as the element closes.
    objects = defaultdict(list)
    paintable = {"ofd:ImageObject", "ofd:PathObject", "ofd:TextObject"}
    owners = {"ofd:Layer", "ofd:Appearance"}

    def collect(parents, key, value):
        if isinstance(value, str) and parents[-1][0] != "ofd:TextCode":
            value = value.strip()
        if key in paintable and isinstance(value, dict):
            draw_param = None
            for parent in reversed(parents[:-1]):
                name, attributes = parent
                if name == "ofd:Clips":
                    break
                if draw_param is None and attributes:
                    draw_param = attributes.get("DrawParam")
                if name in owners:
                    if draw_param is not None:
                        value.setdefault("@DrawParam", draw_param)
                    objects[id(parent)].append((key, value))
                    break
        elif key in owners and isinstance(value, dict):
            value["_objects"] = objects.pop(id(parents[-1]), [])
        return key, value

    parsed = xmltodict.parse(
        _read(archive, path),
        process_namespaces=True,
        namespaces=_NAMESPACE,
        strip_whitespace=False,
        postprocessor=collect,
    )
    if f"ofd:{root}" not in parsed:
        raise ValueError(f"Expected OFD {root} in {path}")
    return parsed[f"ofd:{root}"] or {}


def _page_size(value: Any) -> list[float]:
    box = [float(part) for part in _value(value).split()]
    if len(box) != 4 or not all(map(math.isfinite, box)) or min(box[2:]) <= 0:
        raise ValueError("Invalid OFD physical page box")
    return box


def _resources(archive: ZipFile, common: dict, directory: str) -> tuple[dict, dict, dict]:
    fonts: dict = {}
    drawparams: dict = {}
    images: dict = {}
    for resource_kind in ("ofd:PublicRes", "ofd:DocumentRes", "ofd:PageRes"):
        for location in _items(common.get(resource_kind)):
            member = _path(_value(location), directory)
            resource = _xml(archive, member, "Res")
            base = _path(resource.get("@BaseLoc", "."), posixpath.dirname(member))
            info = PublicResFileParser(resource or {"ofd:Res": {}})()
            for font in info["font"].values():
                if source := _value(font.get("FontFile")):
                    font["FontFile"] = _path(source, base)
            fonts.update(info["font"])
            drawparams.update(info["drawparams"])
            for resource_id, image in DocumentResFileParser(resource or {"ofd:Res": {}})().items():
                image["fileName"] = _path(_value(image["fileName"]), base)
                images[resource_id] = image
    return fonts, drawparams, images


def _image(archive: ZipFile, images: dict, resource_id: str) -> dict:
    if resource_id not in images:
        raise ValueError(f"OFD image resource ID not found: {resource_id}")
    image = images[resource_id]
    if "imgb64" not in image:
        data = _read(archive, image["fileName"])
        suffix = image["suffix"].lower()
        try:
            if suffix in {"jb2", "jbig2"}:
                with fitz.open(stream=data, filetype="jbig2") as image_doc:
                    data = image_doc[0].get_pixmap().tobytes("png")
                image["suffix"] = "png"
            else:
                with Image.open(BytesIO(data)) as source:
                    if suffix in _NATIVE_IMAGES:
                        source.verify()
                    else:
                        with BytesIO() as output:
                            source.convert("RGBA").save(output, format="PNG")
                            data = output.getvalue()
                        image["suffix"] = "png"
        except (OSError, ValueError, RuntimeError) as exc:
            raise ValueError(f"Cannot decode referenced OFD image {image['fileName']} ({suffix})") from exc
        image["imgb64"] = base64.b64encode(data)
    return image


def _annotations(archive: ZipFile, document: dict, directory: str, page_ids: set[str]) -> dict:
    result = defaultdict(list)
    for location in _items(document.get("ofd:Annotations")):
        member = _path(_value(location), directory)
        index = _xml(archive, member, "Annotations")
        for entry in _items(index.get("ofd:Page")):
            page_id = entry.get("@PageID")
            if page_id in page_ids:
                target = _path(_value(entry.get("ofd:FileLoc")), posixpath.dirname(member))
                annotation = _xml(archive, target, "PageAnnot")
                result[page_id].extend(_items(annotation.get("ofd:Annot")))
    return result


def _signatures(archive: ZipFile, body: dict) -> dict:
    result = defaultdict(list)
    location = _value(body.get("ofd:Signatures"))
    if location:
        index_file = _path(location)
        index = _xml(archive, index_file, "Signatures")
        for entry in _items(index.get("ofd:Signature")):
            signature_file = _path(entry.get("@BaseLoc", ""), posixpath.dirname(index_file))
            signature = _xml(archive, signature_file, "Signature")
            signed_info = signature.get("ofd:SignedInfo") or {}
            stamps = _items(signed_info.get("ofd:StampAnnot"))
            if stamps:
                value = _path(
                    _value(signature.get("ofd:SignedValue")) or "SignedValue.dat", posixpath.dirname(signature_file)
                )
                data = _read(archive, value)
                for stamp in stamps:
                    result[stamp.get("@PageRef")].append((stamp, data))
    return result


def _deltas(value: str) -> Iterator[float]:
    tokens = iter(value.split())
    for token in tokens:
        if token == "g":
            count = int(next(tokens, ""))
            delta = float(next(tokens, ""))
            yield from repeat(delta, count)
        else:
            yield float(token)


class _PdfUnicodeFont(TTFont):
    def addObjects(self, doc) -> None:
        super().addObjects(doc)
        # ReportLab 4.4 writes a five/six-digit scalar for non-BMP chars.
        # A PDF ToUnicode destination requires UTF-16BE surrogate pairs.
        identity = self.fontName.removeprefix("DocWenOfd-")
        for name, stream in doc.idToObject.items():
            if isinstance(name, str) and name.startswith("toUnicodeCMap:") and identity in name:
                stream.content = re.sub(
                    r"(?m)^(<[0-9A-F]+>\s+)<([0-9A-F]{5,6})>$",
                    lambda match: f"{match[1]}<{chr(int(match[2], 16)).encode('utf-16-be').hex().upper()}>",
                    stream.content,
                )


class _PdfFonts:
    def __init__(self, archive: ZipFile) -> None:
        self.archive = archive
        self.fonts: dict[str, TTFont | None] = {}
        self.fallback: TTFont | None = None
        self.installed_names = set(pdfmetrics.getRegisteredFontNames())
        self.outlines: dict[str, OutlineFont | None] = {}
        self.payloads: dict[str, bytes] = {}
        self.adapted: dict[tuple[str, tuple[tuple[int, str], ...]], TTFont] = {}

    def _register(self, payload: bytes) -> TTFont:
        identity = uuid4().hex
        font = _PdfUnicodeFont(f"DocWenOfd-{identity}", BytesIO(payload), asciiReadable=False)
        # ReportLab otherwise aliases unrelated embedded subsets with the same
        # PostScript name, even when they have different glyphs for a character.
        font.face.name += f"-{identity}".encode("ascii")
        pdfmetrics.registerFont(font)
        return font

    def _payload(self, source: str) -> bytes:
        if source not in self.payloads:
            self.payloads[source] = _read(self.archive, source)
        return self.payloads[source]

    def embedded(self, resource: dict) -> TTFont | None:
        source = resource.get("FontFile")
        if not source:
            return None
        if source not in self.fonts:
            payload = self._payload(source)
            try:
                self.fonts[source] = self._register(payload)
            except (TTFError, KeyError, ValueError, IndexError, struct.error):
                logger.debug("Cannot embed OFD font %s as TrueType", source, exc_info=True)
                self.fonts[source] = None
        return self.fonts[source]

    def font_for(self, resource: dict, char: str, *, outline: bool = False) -> str:
        candidate = self.embedded(resource)
        if candidate is not None and ord(char) in candidate.face.charToGlyph:
            return candidate.fontName

        installed = (resource.get("FontNameORI"), resource.get("FamilyNameORI"), resource.get("FontName"))
        for name in installed:
            if isinstance(name, str) and name in self.installed_names:
                candidate = pdfmetrics.getFont(name)
                if isinstance(candidate, TTFont) and ord(char) <= 0xFFFF and ord(char) in candidate.face.charToGlyph:
                    return name

        if char.isascii() and not outline:
            family = (resource.get("FamilyName") or resource.get("FontName") or "").lower()
            if family.startswith("times"):
                return "Times-Bold" if resource.get("Bold") in {"true", "1"} else "Times-Roman"
            return "Helvetica-Bold" if resource.get("Bold") in {"true", "1"} else "Helvetica"

        if self.fallback is None:
            self.fallback = self._register(fitz.Font("cjk").buffer)
        if ord(char) not in self.fallback.face.charToGlyph:
            # Preserve the source code point in ToUnicode even if the available
            # fallback font can only show its missing-glyph outline.
            self.fallback.face.charToGlyph[ord(char)] = 0
        return self.fallback.fontName

    def source_outlines(self, resource: dict) -> OutlineFont | None:
        source = resource.get("FontFile")
        if not source:
            return None
        if source not in self.outlines:
            self.outlines[source] = OutlineFont.from_bytes(self._payload(source))
        return self.outlines[source]

    def outline_for(self, font_name: str, codepoint: int) -> Outline | None:
        if font_name not in self.outlines:
            font = pdfmetrics.getFont(font_name)
            self.outlines[font_name] = (
                OutlineFont.from_bytes(cast(Any, font.face)._ttf_data) if isinstance(font, TTFont) else None
            )
        candidate = self.outlines[font_name]
        return candidate.glyph_by_unicode(codepoint) if candidate is not None else None

    def adapted_font(self, resource: dict, source: OutlineFont, mapping: dict[int, str]) -> str:
        key = resource["FontFile"], tuple(sorted(mapping.items()))
        if key not in self.adapted:
            self.adapted[key] = self._register(source.to_truetype(mapping))
        return self.adapted[key].fontName

    def close(self) -> None:
        for font in [*self.fonts.values(), *self.adapted.values(), self.fallback]:
            if font is not None:
                font.unregister()


class _CanvasPathPen(BasePen):
    def __init__(self, path: PDFPathObject, glyphs: Any = None) -> None:
        super().__init__(glyphs)
        self.path = path

    def _moveTo(self, point) -> None:
        self.path.moveTo(*point)

    def _lineTo(self, point) -> None:
        self.path.lineTo(*point)

    def _curveToOne(self, first, second, end) -> None:
        self.path.curveTo(*first, *second, *end)

    def _closePath(self) -> None:
        self.path.close()


def _color(row: dict, key: str, drawparam: dict) -> tuple[float, float, float]:
    color = row.get(f"ofd:{key}") or {}
    channels = color.get("@Value", "").split() or drawparam.get(key, {}).get("value", ())
    values = [float(channel) / 255 for channel in channels if channel]
    if len(values) == 1:
        return values[0], values[0], values[0]
    if len(values) == 3:
        return values[0], values[1], values[2]
    return 0.0, 0.0, 0.0


def _matrix(row: dict, default: str = "1 0 0 1 0 0") -> list[float]:
    return [float(value) for value in row.get("@CTM", default).split()]


def _paint_path(canvas: Canvas, row: dict, size: list[float], scale: float, drawparam: dict) -> None:
    stroke = row.get("@Stroke", "true") in {"true", "1"}
    fill = row.get("@Fill", "false") in {"true", "1"}
    source = _value(row.get("ofd:AbbreviatedData"))
    if (not fill and not stroke) or not source:
        return
    # OFD and SVG share command coordinates. Their cubic, start and close
    # commands use different letters, so adapt those before projecting.
    path = canvas.beginPath()
    parse_path(source.translate(str.maketrans({"S": "M", "B": "C", "C": "Z"})), _CanvasPathPen(path))
    bx, by, _, _ = map(float, row.get("@Boundary", "0 0 0 0").split())
    a, b, c, d, dx, dy = _matrix(row)
    canvas.saveState()
    try:
        canvas.transform(a, -b, c, -d, (bx + dx) * scale, (size[3] - by - dy) * scale)
        canvas.scale(scale, scale)
        canvas.setFillColorRGB(*_color(row, "FillColor", drawparam))
        canvas.setStrokeColorRGB(*_color(row, "StrokeColor", drawparam))
        canvas.setLineWidth(float(row.get("@LineWidth") or drawparam.get("LineWidth") or 0.25))
        if "@Alpha" in row:
            alpha = float(row["@Alpha"]) / 255
            canvas.setFillAlpha(alpha)
            canvas.setStrokeAlpha(alpha)
        fill_rule = 0 if row.get("@Rule") == "Even-Odd" else 1
        canvas.drawPath(path, stroke=int(stroke), fill=int(fill), fillMode=fill_rule)
    finally:
        canvas.restoreState()


def _paint_glyph(
    canvas: Canvas, glyph: Outline, font_size: float, x: float, y: float, angle: float, *, stroke: bool, fill: bool
) -> None:
    units, glyphs, glyph_name = glyph
    path = canvas.beginPath()
    glyphs[glyph_name].draw(_CanvasPathPen(path, glyphs))
    scale = font_size / units
    canvas.saveState()
    try:
        canvas.translate(x, y)
        canvas.rotate(angle)
        canvas.scale(scale, scale)
        canvas.drawPath(path, fill=int(fill), stroke=int(stroke), fillMode=1)
    finally:
        canvas.restoreState()


def _project_glyphs(row: dict, value: str, resource: dict, names: _PdfFonts) -> dict[int, str]:
    if not row.get("ofd:CGTransform") or not resource.get("FontFile"):
        return {}
    embedded = names.embedded(resource)
    cmap = embedded.face.charToGlyph if embedded is not None else {}
    explicit: dict[int, str] = {}
    source = None
    for entry in _items(row["ofd:CGTransform"]):
        start = int(entry["@CodePosition"])
        count = int(entry.get("@CodeCount", 1))
        chars = value[start : start + count]
        gids = [int(number) for number in _deltas(_value(entry.get("ofd:Glyphs")))]
        if (
            len(chars) != count
            or len(gids) != count
            or all(cmap.get(ord(char)) == gid for char, gid in zip(chars, gids, strict=True))
        ):
            continue
        source = names.source_outlines(resource)
        if source is not None:
            for index, glyph in enumerate(source.glyphs_by_id(gids), start):
                if glyph is not None:
                    explicit[index] = glyph[2]
    if source is None or not explicit:
        return {}

    # Each visible subset maps source Unicode to its exact outlines. If an
    # object uses two different outlines for one code point, use a second
    # subset for the conflicting occurrence.
    primary: dict[int, str] = {}
    for index, glyph_name in explicit.items():
        primary.setdefault(ord(value[index]), glyph_name)
    primary_font = names.adapted_font(resource, source, primary)
    projected = {}
    for index, glyph_name in explicit.items():
        codepoint = ord(value[index])
        projected[index] = (
            primary_font
            if primary[codepoint] == glyph_name
            else names.adapted_font(resource, source, {codepoint: glyph_name})
        )
    return projected


def _paint_text(
    canvas: Canvas,
    row: dict,
    fonts: dict,
    names: _PdfFonts,
    size: list[float],
    scale: float,
    drawparams: dict,
    *,
    artifact: bool = False,
) -> None:
    font = fonts.get(row.get("@Font"), {})
    bx, by, _, _ = map(float, row["@Boundary"].split())
    a, b, c, d, dx, dy = _matrix(row)
    font_size = float(row["@Size"]) * scale
    if font_size <= 0:
        return
    paint = drawparams.get(row.get("@DrawParam"), {})
    stroke = row.get("@Stroke", "false") in {"true", "1"}
    fill = row.get("@Fill", "true") in {"true", "1"}
    if not fill and not stroke:
        return
    codes = [{"#text": code} if isinstance(code, str) else code for code in _items(row.get("ofd:TextCode")) if code]
    source_glyphs = _project_glyphs(row, "".join(code.get("#text", "") for code in codes), font, names)

    x = y = 0.0
    direction = -float(row.get("@CharDirection", 0))
    angle = math.radians(direction)
    ca, sa = math.cos(angle), math.sin(angle)
    canvas.saveState()
    try:
        canvas.setFillColorRGB(*_color(row, "FillColor", paint))
        if stroke:
            canvas.setStrokeColorRGB(*_color(row, "StrokeColor", paint))
            canvas.setLineWidth(float(row.get("@LineWidth") or paint.get("LineWidth") or 0.25) * scale)
        if "@Alpha" in row:
            alpha = float(row["@Alpha"]) / 255
            canvas.setFillAlpha(alpha)
            canvas.setStrokeAlpha(alpha)
        canvas.transform(a, -b, -c, d, (bx + dx) * scale, (size[3] - by - dy) * scale)
        text = None if artifact else canvas.beginText()
        previous_font = ""
        index = 0
        for code in codes:
            if not (value := code.get("#text")):
                continue
            x, y = float(code.get("@X", x)), float(code.get("@Y", y))
            px, py = x, y
            x_steps = chain((0.0,), _deltas(code.get("@DeltaX", "")), repeat(0.0))
            y_steps = chain((0.0,), _deltas(code.get("@DeltaY", "")), repeat(0.0))
            for char, step_x, step_y in zip(value, x_steps, y_steps, strict=False):
                px, py = px + step_x, py + step_y
                font_name = source_glyphs.get(index) or names.font_for(font, char, outline=artifact)
                index += 1
                if text is None:
                    if glyph := names.outline_for(font_name, ord(char)):
                        _paint_glyph(
                            canvas, glyph, font_size, px * scale, -py * scale, direction, stroke=stroke, fill=fill
                        )
                else:
                    if font_name != previous_font:
                        text.setFont(font_name, font_size)
                        previous_font = font_name
                    text.setTextRenderMode(2 if stroke and fill else 1 if stroke else 0)
                    text.setTextTransform(ca, sa, -sa, ca, px * scale, -py * scale)
                    text.textOut(char)
        if text is not None:
            canvas.drawText(text)
    finally:
        canvas.restoreState()


def _paint_content(
    archive: ZipFile,
    drawer: DrawPDF,
    canvas: Canvas,
    page: dict,
    resources: tuple[dict, dict, dict],
    names: _PdfFonts,
    size: list[float],
    *,
    artifact: bool = False,
) -> None:
    fonts, drawparams, images = resources
    layers = [page] if artifact else _items((page.get("ofd:Content") or {}).get("ofd:Layer"))
    for layer in layers:
        if not isinstance(layer, dict):
            continue
        for kind, row in layer.get("_objects", ()):
            if kind == "ofd:TextObject":
                _paint_text(canvas, row, fonts, names, size, drawer.OP, drawparams, artifact=artifact)
                continue
            if kind == "ofd:PathObject":
                _paint_path(canvas, row, size, drawer.OP, drawparams.get(row.get("@DrawParam"), {}))
                continue
            projection = cast(dict, ContentFileParser({kind: row})())
            for image in projection["img_list"]:
                _image(archive, images, image["ResourceID"])
            drawer.draw_img(canvas, projection["img_list"], images, size)


def _scoped_resources(
    archive: ZipFile,
    parent: tuple[dict, dict, dict],
    page: dict,
    member: str,
) -> tuple[dict, dict, dict]:
    if not page.get("ofd:PageRes"):
        return parent
    fonts, drawparams, images = parent
    local_fonts, local_params, local_images = _resources(archive, page, posixpath.dirname(member))
    return fonts | local_fonts, drawparams | local_params, images | local_images


def _seal_art(data: bytes, width: float, height: float, check_cancel: Callable[[], None]) -> ImageReader | None:
    extractor = SealExtract()
    try:
        signed, _ = decode(data)
    except PyAsn1Error:
        return None
    parts: list[Any] = []
    extractor.find_octet_strings(signed, parts)
    for part in parts:
        payload = bytes(part)
        if payload.startswith(b"PK\x03\x04"):
            try:
                with ZipFile(BytesIO(payload)) as nested:
                    if "OFD.xml" not in nested.NameToInfo:
                        continue
                    pdf = _render_archive(nested, check_cancel, seal=True)
            except BadZipFile:
                continue
            with fitz.open(stream=pdf, filetype="pdf") as embedded:
                page = embedded[0]
                # Rasterize only the first appearance to at most 216 dpi
                # at its intended size, with a bounded preview dimension.
                sx = min(1200, math.ceil(width * 3)) / page.rect.width
                sy = min(1200, math.ceil(height * 3)) / page.rect.height
                pixels = page.get_pixmap(matrix=fitz.Matrix(sx, sy), alpha=True)
                return ImageReader(BytesIO(pixels.tobytes("png")))
        image = extractor.hex_to_image(payload)
        if image is not None:
            with image:
                return ImageReader(image.copy())
    return None


def _paint_seals(
    canvas: Canvas, stamps: list[tuple[dict, bytes]], size: list[float], scale: float, check_cancel: Callable[[], None]
) -> None:
    for stamp, data in stamps:
        x, y, w, h = map(float, stamp.get("@Boundary", "0 0 0 0").split())
        if w <= 0 or h <= 0:
            continue
        check_cancel()
        art = _seal_art(data, w * scale, h * scale, check_cancel)
        if art is None:
            continue
        canvas.saveState()
        try:
            if clip := stamp.get("@Clip"):
                cx, cy, cw, ch = map(float, clip.split())
                path = canvas.beginPath()
                path.rect((x + cx) * scale, (size[3] - y - cy - ch) * scale, cw * scale, ch * scale)
                canvas.clipPath(path, fill=0, stroke=0)
            canvas.drawImage(art, x * scale, (size[3] - y - h) * scale, w * scale, h * scale, mask="auto")
        finally:
            canvas.restoreState()


def _paint_document(
    archive: ZipFile,
    drawer: DrawPDF,
    canvas: Canvas,
    body: dict,
    names: _PdfFonts,
    check_cancel: Callable[[], None],
    *,
    seal: bool,
) -> int:
    member = _path(_value(body.get("ofd:DocRoot")))
    directory = posixpath.dirname(member)
    document = _xml(archive, member, "Document")
    common = document.get("ofd:CommonData") or {}
    resources = _resources(archive, common, directory)
    pages = _items((document.get("ofd:Pages") or {}).get("ofd:Page"))
    if seal:
        pages = pages[:1]
    templates = {
        entry["@ID"]: _path(entry.get("@BaseLoc", ""), directory) for entry in _items(common.get("ofd:TemplatePage"))
    }
    annotations = _annotations(archive, document, directory, {entry["@ID"] for entry in pages})
    signatures = {} if seal else _signatures(archive, body)

    for entry in pages:
        check_cancel()
        page_member = _path(entry.get("@BaseLoc", ""), directory)
        page = _xml(archive, page_member, "Page")
        layers = defaultdict(list)
        physical_box = (page.get("ofd:Area") or {}).get("ofd:PhysicalBox")
        for link in _items(page.get("ofd:Template")):
            template_member = templates[link["@TemplateID"]]
            template = _xml(archive, template_member, "Page")
            layers[link.get("@ZOrder", "Background")].append((template, template_member))
            if not physical_box:
                physical_box = (template.get("ofd:Area") or {}).get("ofd:PhysicalBox")
        size = _page_size(physical_box or (common.get("ofd:PageArea") or {}).get("ofd:PhysicalBox"))
        canvas.setPageSize((size[2] * drawer.OP, size[3] * drawer.OP))
        canvas.translate(-size[0] * drawer.OP, size[1] * drawer.OP)
        for content, content_member in [*layers["Background"], (page, page_member), *layers["Foreground"]]:
            scoped = _scoped_resources(archive, resources, content, content_member)
            _paint_content(archive, drawer, canvas, content, scoped, names, size)

        _paint_seals(canvas, signatures.get(entry["@ID"], []), size, drawer.OP, check_cancel)
        for annotation in annotations.get(entry["@ID"], ()):
            appearance = annotation.get("ofd:Appearance")
            if annotation.get("@Type") in {"Stamp", "Watermark"} and appearance and "@Boundary" in appearance:
                x, y, width, height = map(float, appearance["@Boundary"].split())
                s = drawer.OP
                canvas.saveState()
                try:
                    clip = canvas.beginPath()
                    clip.rect(x * s, (size[3] - y - height) * s, width * s, height * s)
                    canvas.clipPath(clip, fill=0, stroke=0)
                    canvas.translate(x * s, -y * s)
                    # A repeated decorative stamp belongs to the page appearance.
                    _paint_content(
                        archive,
                        drawer,
                        canvas,
                        appearance,
                        _scoped_resources(archive, resources, page, page_member),
                        names,
                        size,
                        artifact=True,
                    )
                finally:
                    canvas.restoreState()
        canvas.showPage()
    return len(pages)


def _render_archive(archive: ZipFile, check_cancel: Callable[[], None], *, seal: bool = False) -> bytes:
    output = BytesIO()
    canvas = Canvas(output)
    drawer = DrawPDF([{"pdf_name": "document.pdf"}])
    root = _xml(archive, "OFD.xml", "OFD")
    names = _PdfFonts(archive)
    try:
        page_count = 0
        for body in _items(root.get("ofd:DocBody")):
            page_count += _paint_document(archive, drawer, canvas, body, names, check_cancel, seal=seal)
            if seal and page_count:
                break
        if page_count == 0:
            raise ValueError("OFD document declares no pages")
        canvas.save()
    finally:
        names.close()
    return output.getvalue()


def render_ofd_pdf(input_path: str, check_cancel: Callable[[], None]) -> bytes:
    """Read only referenced members and draw each declared page exactly once."""
    with ZipFile(input_path) as archive:
        return _render_archive(archive, check_cancel)
