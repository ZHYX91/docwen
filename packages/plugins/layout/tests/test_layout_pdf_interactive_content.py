from __future__ import annotations

from collections.abc import Iterable
from pathlib import Path
from typing import cast

import pikepdf
import pymupdf
import pytest
from reportlab.pdfgen.canvas import Canvas

from docwen_plugin_layout.operations.converter import PdfMerger, PdfSplitter

from ._layout_conversions_support import _build_fake_context

pytestmark = pytest.mark.contract


def _write_links(path: Path) -> None:
    with pikepdf.Pdf.new() as pdf:
        pages = [pdf.add_blank_page(page_size=(460, 370)) for _ in range(4)]
        for index, (page, rotation) in enumerate(zip(pages, (0, 90, 180, 270), strict=True)):
            page.obj.Rotate = rotation
            if index == 1:
                page.obj.CropBox = pikepdf.Array([30, 40, 410, 330])
            links = []
            for target_index, target in enumerate(pages):
                values = pikepdf.Array([target.obj, pikepdf.Name.XYZ, 82 + target_index, 120, 1.5])
                if target_index == 1:
                    values = pdf.make_indirect(values)
                link = pikepdf.Dictionary(
                    Type=pikepdf.Name.Annot,
                    Subtype=pikepdf.Name.Link,
                    Rect=pikepdf.Array([50, 280 - 30 * target_index, 120, 295 - 30 * target_index]),
                )
                if target_index % 2:
                    link.Dest = values
                else:
                    link.A = pikepdf.Dictionary(S=pikepdf.Name.GoTo, D=values)
                links.append(pdf.make_indirect(link))
            if index == 2:
                layer = pdf.make_indirect(pikepdf.Dictionary(Type=pikepdf.Name.OCG, Name="Layer"))
                links.insert(
                    1,
                    pdf.make_indirect(
                        pikepdf.Dictionary(
                            Type=pikepdf.Name.Annot,
                            Subtype=pikepdf.Name.Link,
                            Rect=pikepdf.Array([180, 170, 215, 200]),
                            A=pikepdf.Dictionary(
                                S=pikepdf.Name.SetOCGState,
                                State=pikepdf.Array([pikepdf.Name.ON, layer]),
                            ),
                        )
                    ),
                )
            page.obj.Annots = pikepdf.Array(links)
        pdf.save(path)

    # pikepdf exposes indirect scalar objects as Python numbers. Use the raw
    # xref APIs to construct a PDF where numbers are indirect on disk.
    with pymupdf.open(path) as document:
        xref, _, _ = document.page_annot_xrefs(2)[0]
        coordinate_xref = document.get_new_xref()
        document.update_object(coordinate_xref, "82")
        zoom_xref = document.get_new_xref()
        document.update_object(zoom_xref, "1.5")
        rect_xref = document.get_new_xref()
        document.update_object(rect_xref, f"[{coordinate_xref} 0 R 280 120 295]")
        target_xref = document.page_xref(0)
        document.xref_set_key(xref, "A/D", f"[{target_xref} 0 R /XYZ {coordinate_xref} 0 R 120 {zoom_xref} 0 R]")
        document.xref_set_key(xref, "Rect", f"{rect_xref} 0 R")
        document.saveIncr()


def _links(path: Path) -> list[list[tuple[int, tuple[float, ...], str, tuple[float, ...]]]]:
    with pikepdf.Pdf.open(path) as pdf:
        page_by_object = {page.obj.objgen: index for index, page in enumerate(pdf.pages)}
        return [
            [
                (
                    page_by_object[target[0].objgen],
                    tuple(float(number) for number in cast(Iterable[float], annotation.Rect)),
                    str(target[1]),
                    tuple(float(target[index]) for index in range(2, len(target))),
                )
                for annotation in cast(Iterable[pikepdf.Object], page.obj.get("/Annots", pikepdf.Array()))
                if "/Dest" in annotation or ("/A" in annotation and str(annotation.A.S) == "/GoTo")
                for target in [annotation.A.D if "/A" in annotation else annotation.Dest]
            ]
            for page in pdf.pages
        ]


def test_pdf_operations_preserve_rotated_link_targets_and_zoom(tmp_path: Path) -> None:
    source = tmp_path / "source.pdf"
    _write_links(source)
    expected = _links(source)
    source_bytes = source.read_bytes()
    leading = tmp_path / "leading.pdf"
    with pikepdf.Pdf.new() as pdf:
        pdf.add_blank_page()
        pdf.save(leading)

    merge = _build_fake_context(str(leading), str(tmp_path), "pdf", action_name="merge_pdfs")
    merge.request.input_refs.append(
        type(merge.request.input_refs[0])(path=str(source), format="pdf", category="layout")
    )
    merged = PdfMerger().convert(merge)
    assert merged.success, merged.error
    assert _links(Path(merged.artifacts[0].staging_path)) == [
        [],
        *[[(page + 1, rect, mode, coords) for page, rect, mode, coords in links] for links in expected],
    ]

    split = _build_fake_context(
        str(source), str(tmp_path), "pdf", {"split_mode": "custom", "pages": [1, 3]}, action_name="split_pdf"
    )
    result = PdfSplitter().convert(split)
    assert result.success, result.error
    for group, artifact in zip(((0, 2), (1, 3)), result.artifacts, strict=True):
        wanted = [
            [(group.index(page), rect, mode, coords) for page, rect, mode, coords in expected[index] if page in group]
            for index in group
        ]
        actual = _links(Path(artifact.staging_path))
        assert [sorted(links) for links in actual] == [sorted(links) for links in wanted]
    assert source.read_bytes() == source_bytes


def test_pdf_split_rejects_cross_page_radio_group_before_staging_any_files(tmp_path: Path) -> None:
    source = tmp_path / "shared-radio.pdf"
    staging = tmp_path / "staging"
    staging.mkdir()
    canvas = Canvas(str(source))
    for page in range(1, 5):
        canvas.drawString(40, 380, f"Distinct page {page}")
        canvas.acroForm.radio(name="delivery", value=str(page), selected=page == 1, x=40, y=340)
        canvas.showPage()
    canvas.save()
    source_bytes = source.read_bytes()

    context = _build_fake_context(
        str(source), str(staging), "pdf", {"split_mode": "custom", "pages": [1, 3]}, action_name="split_pdf"
    )
    result = PdfSplitter().convert(context)

    assert not result.success
    assert result.error is not None
    assert result.error.diagnostic_code == "PDF-SPLIT-UNSUPPORTED-SHARED-FIELDS"
    assert not result.artifacts
    assert not list(staging.iterdir())
    assert source.read_bytes() == source_bytes


@pytest.mark.parametrize(
    ("mode", "groups"),
    [
        ("custom", ((1, 3), (2, 4))),
        ("odd_even", ((1, 3), (2, 4))),
        ("every_page", ((1,), (2,), (3,), (4,))),
    ],
)
def test_pdf_split_preserves_filled_indirect_forms(
    tmp_path: Path, mode: str, groups: tuple[tuple[int, ...], ...]
) -> None:
    source = tmp_path / "forms.pdf"
    canvas = Canvas(str(source))
    for page in range(1, 5):
        canvas.drawString(40, 380, f"Page {page}")
        canvas.acroForm.textfield(name=f"field-{page}", value=f"value-{page}", x=40, y=340, width=120, height=28)
        canvas.showPage()
    canvas.save()
    source_bytes = source.read_bytes()
    with pikepdf.Pdf.open(source) as doc:
        assert doc.Root.AcroForm.is_indirect

    options: dict[str, object] = {"split_mode": mode}
    if mode == "custom":
        options["pages"] = [1, 3]
    context = _build_fake_context(str(source), str(tmp_path), "pdf", options, action_name="split_pdf")
    result = PdfSplitter().convert(context)
    assert result.success, result.error
    for group, artifact in zip(groups, result.artifacts, strict=True):
        with pikepdf.Pdf.open(artifact.staging_path) as doc:
            expected = [(f"field-{page}", f"value-{page}") for page in group]
            fields = cast(Iterable[pikepdf.Object], doc.Root.AcroForm.Fields)
            assert [(str(field.T), str(field.V)) for field in fields] == expected
            assert [
                (str(widget.T), str(widget.V))
                for page in doc.pages
                for widget in cast(Iterable[pikepdf.Object], page.obj.Annots)
                if widget.Subtype == pikepdf.Name.Widget
            ] == expected
    assert source.read_bytes() == source_bytes
