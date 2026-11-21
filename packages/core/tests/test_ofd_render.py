"""Regression checks using small OFD containers generated in the tests."""

from __future__ import annotations

from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest

from docwen_core.ofd import render_ofd_pdf

pytestmark = pytest.mark.unit
_NS = 'xmlns:o="http://www.ofdspec.org/2016"'


def _container(file: Path | BytesIO, members: dict[str, bytes | str]) -> None:
    with ZipFile(file, "w") as archive:
        for path, data in members.items():
            archive.writestr(path, data)


def _ofd(root: str, signatures: str = "") -> str:
    return f"<o:OFD {_NS}><o:DocBody><o:DocRoot>{root}</o:DocRoot>{signatures}</o:DocBody></o:OFD>"


def _document(pages: str, resources: str = "", annotations: str = "") -> str:
    return (
        f"<o:Document {_NS}><o:CommonData>"
        f"<o:PageArea><o:PhysicalBox>0 0 210 297</o:PhysicalBox></o:PageArea>"
        f"{resources}</o:CommonData><o:Pages>{pages}</o:Pages>{annotations}</o:Document>"
    )


def _page(content: str = "", box: str = "") -> str:
    return f"<o:Page {_NS}>{box}<o:Content><o:Layer>{content}</o:Layer></o:Content></o:Page>"


def _text(value: str, *, x: float = 10, y: float = 10, transform: str = "") -> str:
    dx = f' DeltaX="g {len(value) - 1} 4"' if len(value) > 1 else ""
    return (
        f'<o:TextObject ID="7" Font="4" Boundary="{x} {y} 20 10" Size="4" {transform}>'
        f'<o:TextCode X="0" Y="4"{dx}>{value}</o:TextCode></o:TextObject>'
    )


def _png(color: str) -> bytes:
    from PIL import Image

    with Image.new("RGB", (5, 5), color) as image, BytesIO() as output:
        image.save(output, format="PNG")
        return output.getvalue()


def test_pages_follow_declared_order_and_keep_independent_sizes_and_blank_page(tmp_path: Path) -> None:
    fitz = pytest.importorskip("fitz")
    pytest.importorskip("easyofd")
    source = tmp_path / "flat.ofd"
    _container(
        source,
        {
            "OFD.xml": _ofd("Document.xml"),
            "Document.xml": _document(
                '<o:Page ID="47" BaseLoc="Content_9.xml"/><o:Page ID="11" BaseLoc="./Content_1.xml"/>'
                '<o:Page ID="73" BaseLoc="/last/Content_1.xml"/>'
            ),
            "Content_1.xml": _page(_text("Second"), "<o:Area><o:PhysicalBox>0 0 90 60</o:PhysicalBox></o:Area>"),
            "last/Content_1.xml": _page(),
            "Content_9.xml": _page(_text("First")),
        },
    )

    with fitz.open(stream=render_ofd_pdf(str(source)), filetype="pdf") as document:
        assert len(document) == 3
        assert [page.rect.width for page in document] == pytest.approx([595.276, 255.118, 595.276], abs=0.01)
        assert [page.rect.height for page in document] == pytest.approx([841.89, 170.079, 841.89], abs=0.01)
        assert ["".join(page.get_text().split()) for page in document] == ["First", "Second", ""]


@pytest.mark.parametrize("missing", [False, True])
def test_exact_references_ignore_unused_descriptor_and_same_name_resource(tmp_path: Path, missing: bool) -> None:
    fitz = pytest.importorskip("fitz")
    pytest.importorskip("easyofd")
    source = tmp_path / "exact references.ofd"
    members = {
        "~ofd.xml": b"unused",
        "OFD.xml": _ofd("./Doc 42/Document.xml"),
        "junk/Document.xml": b"unused",
        "Doc 42/page/DocumentRes.xml": b"unused",
        "Doc 42/page/picture.png": _png("blue"),
        "Doc 42/Document.xml": _document(
            '<o:Page ID="3" BaseLoc="page/Content.xml"/>',
            "<o:DocumentRes>./DocumentRes.xml</o:DocumentRes>",
        ),
        "Doc 42/DocumentRes.xml": (
            f'<o:Res {_NS} BaseLoc="Images &amp; Media"><o:MultiMedias>'
            '<o:MultiMedia ID="8" Type="Image"><o:MediaFile>picture.png</o:MediaFile></o:MultiMedia>'
            "</o:MultiMedias></o:Res>"
        ),
        "Doc 42/page/Content.xml": _page('<o:ImageObject ID="9" ResourceID="8" Boundary="10 20 30 40"/>'),
    }
    if not missing:
        members["Doc 42/Images & Media/picture.png"] = _png("red")
    _container(source, members)

    if missing:
        with pytest.raises(ValueError, match=r"OFD member not found: Doc 42/Images & Media/picture\.png"):
            render_ofd_pdf(str(source))
    else:
        with fitz.open(stream=render_ofd_pdf(str(source)), filetype="pdf") as document:
            page = document[0]
            assert len(page.get_image_info()) == 1
            assert page.get_pixmap().pixel(50, 80) == (255, 0, 0)


def _font(*, no_cmap: bool = False) -> bytes:
    from fontTools.fontBuilder import FontBuilder
    from fontTools.pens.ttGlyphPen import TTGlyphPen

    builder = FontBuilder(1000, isTTF=True)
    builder.setupGlyphOrder([".notdef", "space", "schematic", "private"])
    builder.setupCharacterMap({32: "space", ord("文"): "schematic", 0xE000: "private"})
    glyphs = {name: TTGlyphPen(None).glyph() for name in [".notdef", "space"]}
    for name, w in {"schematic": 650, "private": 400}.items():
        pen = TTGlyphPen(None)
        pen.moveTo((0, 0))
        pen.lineTo((w, 0))
        pen.lineTo((w, 700))
        pen.lineTo((0, 700))
        pen.closePath()
        glyphs[name] = pen.glyph()
    builder.setupGlyf(glyphs)
    builder.setupHorizontalMetrics(dict.fromkeys(glyphs, (1000, 0)))
    builder.setupHorizontalHeader(ascent=800, descent=-200)
    builder.setupNameTable({"familyName": "Synthetic", "styleName": "Regular", "psName": "Synthetic-Regular"})
    builder.setupOS2()
    builder.setupPost()
    if no_cmap:
        del builder.font["cmap"]
    with BytesIO() as output:
        builder.save(output)
        return output.getvalue()


@pytest.mark.parametrize("no_cmap", [False, True])
def test_embedded_text_has_visual_glyphs_and_no_duplicate_stamp_text(tmp_path: Path, no_cmap: bool) -> None:
    fitz = pytest.importorskip("fitz")
    pytest.importorskip("easyofd")
    source = tmp_path / "embedded font.ofd"
    body = (
        '<o:TextObject ID="7" Font="4" Boundary="10 10 20 10" Size="4">'
        '<o:CGTransform CodePosition="0" CodeCount="2" GlyphCount="2"><o:Glyphs>2 3</o:Glyphs></o:CGTransform>'
        '<o:TextCode X="0" Y="4" DeltaX="4">文&#xE000;</o:TextCode></o:TextObject>'
    )
    stamp = _text("文", x=0, y=0, transform='Alpha="51"').replace(
        "<o:TextCode", '<o:CGTransform CodePosition="0"><o:Glyphs>2</o:Glyphs></o:CGTransform><o:TextCode'
    )
    _container(
        source,
        {
            "OFD.xml": _ofd("Doc/Document.xml"),
            "Doc/Document.xml": _document(
                '<o:Page ID="11" BaseLoc="odd/Page_9.xml"/>',
                "<o:PublicRes>fonts.xml</o:PublicRes>",
                "<o:Annotations>notes/all.xml</o:Annotations>",
            ),
            "Doc/fonts.xml": (
                f'<o:Res {_NS} BaseLoc="fonts"><o:Fonts><o:Font ID="4" FontName="Synthetic">'
                "<o:FontFile>glyphs.ttf</o:FontFile></o:Font></o:Fonts></o:Res>"
            ),
            "Doc/fonts/glyphs.ttf": _font(no_cmap=no_cmap),
            "Doc/odd/Page_9.xml": _page(body),
            "Doc/notes/all.xml": (
                f'<o:Annotations {_NS}><o:Page PageID="11"><o:FileLoc>note.xml</o:FileLoc></o:Page></o:Annotations>'
            ),
            "Doc/notes/note.xml": (
                f'<o:PageAnnot {_NS}><o:Annot ID="12" Type="Stamp"><o:Appearance Boundary="70 40 40 20">'
                f"{stamp}</o:Appearance></o:Annot></o:PageAnnot>"
            ),
        },
    )

    with fitz.open(stream=render_ofd_pdf(str(source)), filetype="pdf") as document:
        page = document[0]
        assert "".join(page.get_text().split()) == "文\ue000"
        assert all(span["type"] == 0 for span in page.get_texttrace())
        pixmap = page.get_pixmap()
        assert pixmap.pixel(30, 35) == (0, 0, 0)
        assert pixmap.pixel(41, 35) == (0, 0, 0)
        assert pixmap.pixel(200, 120) == (204, 204, 204)


@pytest.mark.parametrize("nested", [False, True])
def test_relative_signature_reference_paints_seal_on_referenced_page(tmp_path: Path, nested: bool) -> None:
    fitz = pytest.importorskip("fitz")
    pytest.importorskip("easyofd")
    from pyasn1.codec.der.encoder import encode
    from pyasn1.type.univ import OctetString

    payload = _png("red")
    if nested:
        with BytesIO() as seal:
            _container(
                seal,
                {
                    "OFD.xml": _ofd("Document.xml", "<o:Signatures>intentionally absent.xml</o:Signatures>"),
                    "Document.xml": _document(
                        '<o:Page ID="20" BaseLoc="Image.xml"/><o:Page ID="21" BaseLoc="missing page.xml"/>'
                    ),
                    "Image.xml": _page(
                        '<o:PathObject ID="22" Boundary="0 0 20 20" Fill="true" Stroke="false">'
                        '<o:FillColor Value="255 0 0"/><o:AbbreviatedData>M 0 0 L 20 0 L 20 20 L 0 20 C</o:AbbreviatedData>'
                        "</o:PathObject>",
                        "<o:Area><o:PhysicalBox>0 0 20 20</o:PhysicalBox></o:Area>",
                    ),
                },
            )
            payload = seal.getvalue()
    source = tmp_path / "seal.ofd"
    _container(
        source,
        {
            "OFD.xml": _ofd("Doc/Document.xml", "<o:Signatures>Doc/signs/all.xml</o:Signatures>"),
            "Doc/Document.xml": _document(
                '<o:Page ID="11" BaseLoc="Content_2.xml"/><o:Page ID="12" BaseLoc="Content_1.xml"/>'
            ),
            "Doc/Content_2.xml": _page(),
            "Doc/Content_1.xml": _page(),
            "Doc/signs/all.xml": f'<o:Signatures {_NS}><o:Signature ID="7" BaseLoc="stamp/Signature.xml"/></o:Signatures>',
            "Doc/signs/stamp/Signature.xml": (
                f'<o:Signature {_NS}><o:SignedInfo><o:StampAnnot ID="8" PageRef="12" Boundary="30 30 25 25" Clip="10 0 15 25"/>'
                "</o:SignedInfo><o:SignedValue>seal.dat</o:SignedValue></o:Signature>"
            ),
            "Doc/signs/stamp/seal.dat": encode(OctetString(payload)),
        },
    )

    with fitz.open(stream=render_ofd_pdf(str(source)), filetype="pdf") as document:
        assert [len(page.get_image_info()) for page in document] == [0, 1]
        pixmap = document[1].get_pixmap()
        assert pixmap.pixel(100, 100) == (255, 255, 255)
        assert pixmap.pixel(130, 100) == (255, 0, 0)


def test_page_resources_transformed_background_and_supplementary_text(tmp_path: Path) -> None:
    fitz = pytest.importorskip("fitz")
    pytest.importorskip("easyofd")
    source = tmp_path / "page resources.ofd"
    background = (
        '<o:PathObject ID="8" Boundary="0 0 20 20" CTM="1 0 0 1 8 8" Fill="true" Stroke="false">'
        '<o:FillColor Value="255 0 0"/><o:AbbreviatedData>M 0 0 L 20 0 L 20 20 L 0 20 C</o:AbbreviatedData>'
        "</o:PathObject>"
    )
    page = _page(
        background + _text("文𠀀"),
        "<o:PageRes>./own/fonts.xml</o:PageRes>",
    )
    _container(
        source,
        {
            "OFD.xml": _ofd("Document.xml"),
            "Document.xml": _document('<o:Page ID="1" BaseLoc="Page/Content.xml"/>'),
            "Page/Content.xml": page,
            "Page/own/fonts.xml": (
                f'<o:Res {_NS}><o:Fonts><o:Font ID="4" FontName="Synthetic">'
                "<o:FontFile>glyphs.ttf</o:FontFile></o:Font></o:Fonts></o:Res>"
            ),
            "Page/own/glyphs.ttf": _font(),
        },
    )

    with fitz.open(stream=render_ofd_pdf(str(source)), filetype="pdf") as document:
        page = document[0]
        assert "".join(page.get_text().split()) == "文𠀀"
        assert any("Synthetic-Regular" in font[3] for font in page.get_fonts())
        pixmap = page.get_pixmap()
        assert pixmap.pixel(4, 4) == (255, 255, 255)
        assert pixmap.pixel(25, 25) == (255, 0, 0)
        assert pixmap.pixel(30, 35) == (0, 0, 0)
