from __future__ import annotations

from pathlib import Path

import pikepdf
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
            page.obj.Annots = pikepdf.Array(links)
        pdf.save(path)


def _links(path: Path) -> list[list[tuple[int, tuple[float, ...], str, tuple[float, ...]]]]:
    with pikepdf.Pdf.open(path) as pdf:
        page_by_object = {page.obj.objgen: index for index, page in enumerate(pdf.pages)}
        return [
            [
                (
                    page_by_object[target[0].objgen],
                    tuple(float(number) for number in annotation.Rect),
                    str(target[1]),
                    tuple(float(number) for number in list(target)[2:]),
                )
                for annotation in page.obj.get("/Annots", [])
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
            assert [(str(field.T), str(field.V)) for field in doc.Root.AcroForm.Fields] == expected
            assert [
                (str(widget.T), str(widget.V))
                for page in doc.pages
                for widget in page.obj.Annots
                if widget.Subtype == pikepdf.Name.Widget
            ] == expected
    assert source.read_bytes() == source_bytes
