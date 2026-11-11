"""Authored objects coexist with original unsupported drawing and note data."""

from __future__ import annotations

from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest
from lxml import etree
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image as XlsxImage
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.table import Table
from PIL import Image

from docwen_plugin_markdown.xlsx_template import (
    load_template_workbook,
    register_template_clone,
    save_xlsx_workbook,
)
from docwen_plugin_markdown.xlsx_template.objects import merge_sheet_objects
from docwen_plugin_markdown.xlsx_template.package import (
    PackageEditor,
    parse_xml,
    relocate_formula_references,
    resolve_target,
    serialize_xml,
)

pytestmark = pytest.mark.contract

_S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_D = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
_V = "urn:schemas-microsoft-com:vml"
_X = "urn:schemas-microsoft-com:office:excel"
_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_SHEET = "xl/worksheets/sheet1.xml"


def _package(workbook) -> dict[str, bytes]:
    stream = BytesIO()
    workbook.save(stream)
    with ZipFile(stream) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def _bytes(parts: dict[str, bytes]) -> BytesIO:
    stream = BytesIO()
    with ZipFile(stream, "w") as archive:
        for name, payload in parts.items():
            archive.writestr(name, payload)
    stream.seek(0)
    return stream


def _target(parts: dict[str, bytes], owner: str, suffix: str) -> str:
    parent, name = owner.rsplit("/", 1)
    relations = parse_xml(parts[f"{parent}/_rels/{name}.rels"])
    relation = next(node for node in relations if node.get("Type", "").endswith("/" + suffix))
    return resolve_target(owner, relation.get("Target", ""))


def _merge(source, donor, worksheet, *, images=0, comments=(), hyperlinks=()):
    editor = PackageEditor(source)
    root = parse_xml(source[_SHEET])
    changed = merge_sheet_objects(
        editor,
        _SHEET,
        root,
        donor,
        _SHEET,
        parse_xml(donor[_SHEET]),
        worksheet,
        images,
        set(comments),
        set(hyperlinks),
    )
    if changed:
        editor.parts[_SHEET] = serialize_xml(root)
    editor.finalize_content_types()
    return editor, root


def test_appending_picture_preserves_original_shapes_and_media(tmp_path: Path) -> None:
    old_image = tmp_path / "old.png"
    new_image = tmp_path / "new.png"
    Image.new("RGB", (6, 7), "red").save(old_image)
    Image.new("RGB", (8, 9), "blue").save(new_image)
    workbook = Workbook()
    sheet = workbook.worksheets[0]
    sheet.add_image(XlsxImage(str(old_image)), "B2")
    source = _package(workbook)
    donor_workbook = load_workbook(_bytes(source))
    donor_sheet = donor_workbook.worksheets[0]
    donor_sheet.add_image(XlsxImage(str(new_image)), "D4")
    donor = _package(donor_workbook)
    drawing_part = _target(source, _SHEET, "drawing")
    drawing = parse_xml(source[drawing_part])
    original_anchor = serialize_xml(drawing[0])
    opaque = etree.SubElement(drawing, f"{{{_D}}}absoluteAnchor")
    shape = etree.SubElement(opaque, f"{{{_D}}}sp")
    shape.set("{urn:test:original-shape}preserve", "yes")
    properties = etree.SubElement(shape, f"{{{_D}}}nvSpPr")
    etree.SubElement(properties, f"{{{_D}}}cNvPr", id="999", name="Original unsupported shape")
    opaque_bytes = serialize_xml(opaque)
    source[drawing_part] = serialize_xml(drawing)
    media_before = {name: data for name, data in source.items() if name.startswith("xl/media/")}

    editor, _root = _merge(source, donor, donor_sheet, images=1)

    after = parse_xml(editor.parts[drawing_part])
    assert len(after) == 3
    assert serialize_xml(after[0]) == original_anchor
    assert serialize_xml(after[1]) == opaque_bytes
    assert all(editor.parts[name] == data for name, data in media_before.items())
    ids = [node.get("id") for node in after.iter(f"{{{_D}}}cNvPr")]
    assert len(ids) == len(set(ids))
    assert int(ids[-1]) > 999
    assert new_image.read_bytes() in editor.parts.values()
    for node in after[-1].iter():
        for key, value in node.attrib.items():
            if key.startswith("{" + _R + "}"):
                target, _kind, mode = editor.relationship_target(drawing_part, value)
                assert mode == "External" or target in editor.parts


def _note_shape(root, row: int, col: int):
    return next(
        shape
        for shape in root.iter(f"{{{_V}}}shape")
        if shape.findtext(f"{{{_X}}}ClientData/{{{_X}}}Row") == str(row)
        and shape.findtext(f"{{{_X}}}ClientData/{{{_X}}}Column") == str(col)
    )


def test_changed_notes_keep_other_rich_text_and_vml_geometry() -> None:
    workbook = Workbook()
    sheet = workbook.worksheets[0]
    sheet["A1"].comment = Comment("Original A", "First")
    sheet["C3"].comment = Comment("Unchanged C", "Second")
    sheet["D4"].comment = Comment("Removed D", "First")
    source = _package(workbook)
    donor_workbook = load_workbook(_bytes(source))
    donor_sheet = donor_workbook.worksheets[0]
    donor_sheet["A1"].comment = Comment("Changed A", "New Author")
    donor_sheet["B2"].comment = Comment("Added B", "New Author")
    donor_sheet["D4"].comment = None
    donor = _package(donor_workbook)
    comments_part = _target(source, _SHEET, "comments")
    comments = parse_xml(source[comments_part])
    untouched = comments.find(f"{{{_S}}}commentList/{{{_S}}}comment[@ref='C3']")
    text = untouched.find(f"{{{_S}}}text")
    text.clear()
    run = etree.SubElement(text, f"{{{_S}}}r")
    properties = etree.SubElement(run, f"{{{_S}}}rPr")
    etree.SubElement(properties, f"{{{_S}}}b")
    etree.SubElement(run, f"{{{_S}}}t").text = "Unchanged C"
    untouched.set("{urn:test:note}metadata", "keep")
    expected_comment = serialize_xml(untouched)
    source[comments_part] = serialize_xml(comments)
    vml_part = _target(source, _SHEET, "vmlDrawing")
    vml = parse_xml(source[vml_part])
    original_a = _note_shape(vml, 0, 0)
    original_a.set("style", "position:absolute;width:407pt;height:193pt;visibility:visible")
    original_c = _note_shape(vml, 2, 2)
    original_c.set("{urn:test:note}metadata", "keep")
    expected_a, expected_c = serialize_xml(original_a), serialize_xml(original_c)
    source[vml_part] = serialize_xml(vml)

    editor, _root = _merge(source, donor, donor_sheet, comments={"A1", "B2", "D4"})

    result = parse_xml(editor.parts[comments_part])
    note_map = {node.get("ref"): node for node in result.find(f"{{{_S}}}commentList")}
    assert set(note_map) == {"A1", "B2", "C3"}
    assert serialize_xml(note_map["C3"]) == expected_comment
    assert "".join(note_map["A1"].find(f"{{{_S}}}text").itertext()) == "Changed A"
    authors = [node.text for node in result.find(f"{{{_S}}}authors")]
    assert authors[int(note_map["A1"].get("authorId"))] == "New Author"
    shapes = parse_xml(editor.parts[vml_part])
    assert serialize_xml(_note_shape(shapes, 0, 0)) == expected_a
    assert serialize_xml(_note_shape(shapes, 2, 2)) == expected_c
    assert _note_shape(shapes, 1, 1) is not None
    assert len(list(shapes.iter(f"{{{_V}}}shape"))) == 3
    ids = [node.get("id") for node in shapes.iter(f"{{{_V}}}shape")]
    assert len(ids) == len(set(ids))


def test_hyperlink_changes_preserve_other_relations_and_extension() -> None:
    workbook = Workbook()
    sheet = workbook.worksheets[0]
    sheet["A1"] = "Changed link"
    sheet["A1"].hyperlink = "https://example.com/old"
    sheet["C3"] = "Untouched link"
    sheet["C3"].hyperlink = "https://example.com/untouched"
    sheet["D4"] = "Removed link"
    sheet["D4"].hyperlink = "https://example.com/remove"
    source = _package(workbook)
    donor_workbook = load_workbook(_bytes(source))
    donor_sheet = donor_workbook.worksheets[0]
    donor_sheet["A1"].hyperlink = "https://example.com/new"
    donor_sheet["B2"].hyperlink = Hyperlink(ref="B2", location="Sheet!A1", tooltip="internal")
    donor_sheet["D4"].hyperlink = None
    donor = _package(donor_workbook)
    root = parse_xml(source[_SHEET])
    untouched = root.find(f"{{{_S}}}hyperlinks/{{{_S}}}hyperlink[@ref='C3']")
    untouched.set("{urn:test:link}metadata", "preserve")
    expected = serialize_xml(untouched)
    extension = etree.SubElement(root, f"{{{_S}}}extLst")
    etree.SubElement(extension, f"{{{_S}}}ext", uri="urn:test:extension")
    expected_ext = serialize_xml(extension)
    source[_SHEET] = serialize_xml(root)

    editor, root = _merge(source, donor, donor_sheet, hyperlinks={"A1", "B2", "D4"})

    links = {node.get("ref"): node for node in root.find(f"{{{_S}}}hyperlinks")}
    assert set(links) == {"A1", "B2", "C3"}
    assert serialize_xml(links["C3"]) == expected
    assert serialize_xml(root.find(f"{{{_S}}}extLst")) == expected_ext
    target, _kind, mode = editor.relationship_target(_SHEET, links["A1"].get(f"{{{_R}}}id"))
    assert target == "https://example.com/new" and mode == "External"
    assert links["B2"].get("location") == "Sheet!A1"


def test_new_comment_and_picture_carriers_precede_existing_extensions(tmp_path: Path) -> None:
    workbook = Workbook()
    source = _package(workbook)
    root = parse_xml(source[_SHEET])
    etree.SubElement(root, f"{{{_S}}}extLst")
    source[_SHEET] = serialize_xml(root)
    image = tmp_path / "new.png"
    Image.new("RGB", (3, 3), "green").save(image)
    sheet = workbook.worksheets[0]
    sheet["B2"].comment = Comment("New", "Author")
    sheet.add_image(XlsxImage(str(image)), "A1")
    donor = _package(workbook)

    editor, after = _merge(source, donor, sheet, comments={"B2"})

    names = [etree.QName(node).localname for node in after]
    assert names.index("drawing") < names.index("legacyDrawing") < names.index("extLst")
    reopened = load_workbook(_bytes(editor.parts))
    note = reopened.worksheets[0]["B2"].comment
    assert note is not None and note.text == "New"
    images = vars(reopened.worksheets[0]).get("_images")
    assert isinstance(images, list) and len(images) == 1


def test_unchanged_objects_leave_every_part_byte_identical() -> None:
    workbook = Workbook()
    workbook.worksheets[0]["A1"].comment = Comment("Original", "Author")
    workbook.worksheets[0]["A1"].hyperlink = "https://example.com/keep"
    source = _package(workbook)

    editor, _root = _merge(source, source, workbook.worksheets[0])

    assert editor.parts == source


def test_clone_graph_renames_tables_and_keeps_source_parts_untouched() -> None:
    workbook = Workbook()
    sheet = workbook.worksheets[0]
    sheet.append(["Label", "Value"])
    sheet.append(["first", 4])
    sheet.add_table(Table(displayName="Items", ref="A1:B2"))
    workbook.defined_names.add(DefinedName("Items_copy1", attr_text="Sheet!$B$2"))
    sheet["D2"] = '=SUM(Items[Value])+SUM(Items)+LEN("Items[Value]")'
    source = _package(workbook)
    source_root = parse_xml(source[_SHEET])
    formula_cell = source_root.find(f"{{{_S}}}sheetData/{{{_S}}}row/{{{_S}}}c[@r='D2']")
    formula_cell.find(f"{{{_S}}}v").text = "73"
    source[_SHEET] = serialize_xml(source_root)
    source_table = _target(source, _SHEET, "table")
    editor = PackageEditor(source)

    cloned_sheet = editor.clone_part(_SHEET).root_part
    editor.finalize_content_types()

    cloned_table = _target(editor.parts, cloned_sheet, "table")
    table = parse_xml(editor.parts[cloned_table])
    assert table.get("displayName") == "Items_copy2"
    assert table.get("name") == "Items_copy2"
    assert table.get("id") != parse_xml(source[source_table]).get("id")
    clone_root = parse_xml(editor.parts[cloned_sheet])
    cloned_cell = clone_root.find(f"{{{_S}}}sheetData/{{{_S}}}row/{{{_S}}}c[@r='D2']")
    assert cloned_cell.findtext(f"{{{_S}}}f") == 'SUM(Items_copy2[Value])+SUM(Items_copy2)+LEN("Items[Value]")'
    assert cloned_cell.find(f"{{{_S}}}v") is None
    assert all(editor.parts[name] == payload for name, payload in source.items() if name != "[Content_Types].xml")


def test_png_deduplication_only_removes_new_parts_and_rewrites_all_relations(tmp_path: Path) -> None:
    image = tmp_path / "image.png"
    Image.new("RGB", (3, 4), "orange").save(image)
    workbook = Workbook()
    workbook.worksheets[0].add_image(XlsxImage(str(image)), "A1")
    workbook.worksheets[0].add_image(XlsxImage(str(image)), "C3")
    source = _package(workbook)
    originals = {name for name in source if name.startswith("xl/media/")}
    assert len(originals) == 2
    editor = PackageEditor(source)
    donor_part = sorted(originals)[0]
    new_part = editor.import_part(source, donor_part)
    link = editor.add_relationship(_SHEET, _R + "/image", new_part)

    editor.deduplicate_new_png_media(set(source))
    editor.finalize_content_types()

    assert all(editor.parts[name] == source[name] for name in originals)
    assert new_part not in editor.parts
    destination, _kind, _mode = editor.relationship_target(_SHEET, link)
    assert destination in originals
    assert all(name in editor.parts for name in source)


@pytest.mark.parametrize("target", ["../../../escape.xml", "https://example.com/invalid", "..\\invalid.xml"])
def test_imported_relationship_graph_rejects_non_package_targets(target: str) -> None:
    source = _package(Workbook())
    relationship_namespace = "http://schemas.openxmlformats.org/package/2006/relationships"
    relations = etree.Element(f"{{{relationship_namespace}}}Relationships")
    etree.SubElement(
        relations,
        f"{{{relationship_namespace}}}Relationship",
        Id="rId1",
        Type="urn:test:unknown",
        Target=target,
    )
    source["xl/worksheets/_rels/sheet1.xml.rels"] = serialize_xml(relations)
    editor = PackageEditor(_package(Workbook()))

    with pytest.raises(ValueError, match=r"relationship|outside"):
        editor.import_part(source, _SHEET)


@pytest.mark.parametrize("note_cell", ["A1", "B2"])
def test_existing_threaded_comment_is_preserved_and_same_cell_note_is_refused(note_cell: str) -> None:
    workbook = Workbook()
    source = _package(workbook)
    threaded_namespace = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
    threaded = parse_xml(f'<ThreadedComments xmlns="{threaded_namespace}"/>'.encode())
    original = etree.SubElement(
        threaded,
        f"{{{threaded_namespace}}}threadedComment",
        ref="A1",
        personId="{00000000-0000-0000-0000-000000000001}",
        id="{00000000-0000-0000-0000-000000000002}",
    )
    etree.SubElement(original, f"{{{threaded_namespace}}}text").text = "Keep the original thread"
    threaded_part = "xl/threadedComments/threadedComment1.xml"
    editor = PackageEditor(source)
    editor.parts[threaded_part] = serialize_xml(threaded)
    editor.ensure_content_type(threaded_part, "application/vnd.ms-excel.threadedcomments+xml")
    editor.add_relationship(
        _SHEET, "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment", threaded_part
    )
    editor.finalize_content_types()
    source = dict(editor.parts)
    workbook.worksheets[0][note_cell].comment = Comment("Authored classic note", "Author")
    donor = _package(workbook)

    if note_cell == "A1":
        with pytest.raises(ValueError, match=r"threaded comment: A1"):
            merge_sheet_objects(
                editor,
                _SHEET,
                parse_xml(source[_SHEET]),
                donor,
                _SHEET,
                parse_xml(donor[_SHEET]),
                workbook.worksheets[0],
                0,
                {note_cell},
                set(),
            )
        assert editor.parts == source
    else:
        updated, _root = _merge(source, donor, workbook.worksheets[0], comments={note_cell})
        assert updated.parts[threaded_part] == source[threaded_part]
        assert _target(updated.parts, _SHEET, "threadedComment") == threaded_part
        assert _target(updated.parts, _SHEET, "comments") in updated.parts


def test_cloned_chart_and_worksheet_reference_the_filled_sheet_only() -> None:
    workbook = Workbook()
    sheet = workbook.worksheets[0]
    sheet.title = "O'Brien Source"
    sheet.append(["Label", "Value"])
    sheet.append(["Item", 4])
    sheet["D2"] = "=SUM('O''Brien Source'!B2)+LEN(\"O'Brien Source!B2\")"
    chart = BarChart()
    chart.add_data(Reference(sheet, min_col=2, min_row=1, max_row=2), titles_from_data=True)
    chart.set_categories(Reference(sheet, min_col=1, min_row=2, max_row=2))
    sheet.add_chart(chart, "F1")
    source = _package(workbook)
    original_drawing = _target(source, _SHEET, "drawing")
    original_chart = _target(source, original_drawing, "chart")
    chart_root = parse_xml(source[original_chart])
    for reference in chart_root.findall(f".//{{{_C}}}numRef"):
        cache = etree.SubElement(reference, f"{{{_C}}}numCache")
        etree.SubElement(cache, f"{{{_C}}}formatCode").text = "General"
        etree.SubElement(cache, f"{{{_C}}}ptCount", val="1")
        point = etree.SubElement(cache, f"{{{_C}}}pt", idx="0")
        etree.SubElement(point, f"{{{_C}}}v").text = "4"
    source[original_chart] = serialize_xml(chart_root)
    editor = PackageEditor(source)

    cloned_sheet = editor.clone_part(
        _SHEET, source_sheet_name="O'Brien Source", target_sheet_name="Filled report"
    ).root_part
    editor.finalize_content_types()

    clone_drawing = _target(editor.parts, cloned_sheet, "drawing")
    clone_chart = _target(editor.parts, clone_drawing, "chart")
    cloned_chart = parse_xml(editor.parts[clone_chart])
    formula_texts = [node.text for node in cloned_chart.iter(f"{{{_C}}}f")]
    assert formula_texts and all(text is not None and text.startswith("'Filled report'!") for text in formula_texts)
    assert not cloned_chart.findall(f".//{{{_C}}}numCache")
    assert not cloned_chart.findall(f".//{{{_C}}}strCache")
    cloned_root = parse_xml(editor.parts[cloned_sheet])
    formula = cloned_root.findtext(f"{{{_S}}}sheetData/{{{_S}}}row/{{{_S}}}c[@r='D2']/{{{_S}}}f")
    assert formula == "SUM('Filled report'!B2)+LEN(\"O'Brien Source!B2\")"
    assert all(editor.parts[name] == payload for name, payload in source.items() if name != "[Content_Types].xml")


@pytest.mark.parametrize("equals_prefix", ["", "="])
def test_formula_relocation_leaves_external_3d_and_string_references(equals_prefix: str) -> None:
    body = (
        "SUM('O''Brien Source'!$B$2,'[Book.xlsx]O''Brien Source'!$B$2,"
        "'O''Brien Source:Other'!$B$2)+LEN(\"O'Brien Source!B2\")"
    )

    result = relocate_formula_references(equals_prefix + body, "O'Brien Source", "Filled report")

    assert result == equals_prefix + body.replace("'O''Brien Source'!$B$2", "'Filled report'!$B$2", 1)


@pytest.mark.parametrize(
    ("formula", "expected"),
    [
        ("='Prototype'!A1#", "='Filled report'!A1#"),
        ("SUM(Prototype!A1#)", "SUM('Filled report'!A1#)"),
        ("=@Prototype!A1", "=@'Filled report'!A1"),
        ("=@'Prototype'!A1#", "=@'Filled report'!A1#"),
        ("=IFERROR(#N/A/Prototype!B1,0)", "=IFERROR(#N/A/'Filled report'!B1,0)"),
        (
            "=Prototype!LocalLambda(1)+SUM('Prototype'!LocalLambda(2))",
            "='Filled report'!LocalLambda(1)+SUM('Filled report'!LocalLambda(2))",
        ),
        (
            "=SUM(Items[Value], 'Prototype'!A1#) + LEN(\"Items[Value] Prototype!A1#\")",
            "=SUM(Items[Value], 'Filled report'!A1#) + LEN(\"Items[Value] Prototype!A1#\")",
        ),
    ],
)
def test_formula_relocation_accepts_spill_and_implicit_intersection(formula: str, expected: str) -> None:
    assert relocate_formula_references(formula, "Prototype", "Filled report") == expected


def test_formula_relocation_preserves_lexical_boundaries_and_original_spacing() -> None:
    formula = (
        "=SUM(  Prototype!A1#,\nPrototypeExtra!A1#,OtherPrototype!A1#,"
        "'[Book.xlsx]Prototype'!A1#,Prototype:Other!A1,Other:Prototype!A1,"
        'Items[Prototype!A1]) + LEN("Prototype!A1# and ""quoted""")'
    )

    relocated = relocate_formula_references(formula, "Prototype", "Filled report")

    assert relocated == formula.replace("  Prototype!A1#", "  'Filled report'!A1#", 1)


def test_formula_relocation_keeps_escaped_brackets_in_structured_headers() -> None:
    # The column name is x]Prototype!A1[tail; ' escapes each literal bracket.
    structured = "Items[[#All],[x']Prototype!A1'[tail]]"
    formula = f"=SUM({structured}, Prototype!A1#)"

    assert relocate_formula_references(formula, "Prototype", "Filled report") == (
        f"=SUM({structured}, 'Filled report'!A1#)"
    )


@pytest.mark.parametrize(
    ("formula", "expected"),
    [("=#REF!", "=#REF!"), ("SUM(#REF!,ReF!A1)", "SUM(#REF!,'Filled report'!A1)")],
)
def test_formula_relocation_distinguishes_error_literal_from_sheet_name(formula: str, expected: str) -> None:
    assert relocate_formula_references(formula, "ReF", "Filled report") == expected


@pytest.mark.parametrize("table_name", ["Items", "N"])
def test_public_clone_save_relocates_spill_formula_and_renames_table(tmp_path: Path, table_name: str) -> None:
    template = tmp_path / "template.xlsx"
    output = tmp_path / "filled.xlsx"
    workbook = Workbook()
    source = workbook.worksheets[0]
    source.title = "Prototype"
    source.append(["Label", "Value"])
    source.append(["Item", 4])
    source.add_table(Table(displayName=table_name, ref="A1:B2"))
    source["D2"] = "=_xlfn.SEQUENCE(2)"
    original_formula = (
        f"=SUM({table_name}[Value])+SUM('Prototype'!D2#)+SUM(@Prototype!D2)"
        f'+LEN("{table_name}[Value] Prototype!D2#")+IFERROR(#N/A,0)+IFERROR(#REF!/{table_name},0)'
    )
    source["E2"] = original_formula
    workbook.save(template)
    with ZipFile(template) as package:
        original_sheet = package.read(_SHEET)
        original_table = package.read("xl/tables/table1.xml")

    loaded = load_template_workbook(template)
    prototype = loaded["Prototype"]
    clone = loaded.copy_worksheet(prototype)
    clone.title = "Filled report"
    register_template_clone(loaded, clone, prototype)
    clone["B2"] = 99
    save_xlsx_workbook(loaded, output)

    reopened = load_workbook(output)
    assert reopened["Prototype"]["E2"].value == original_formula
    filled = reopened["Filled report"]
    assert filled["B2"].value == 99
    assert filled["D2"].value == "=_xlfn.SEQUENCE(2)"
    assert set(filled.tables) == {f"{table_name}_copy1"}
    assert filled["E2"].value == (
        f"=SUM({table_name}_copy1[Value])+SUM('Filled report'!D2#)+SUM(@'Filled report'!D2)"
        f'+LEN("{table_name}[Value] Prototype!D2#")+IFERROR(#N/A,0)+IFERROR(#REF!/{table_name}_copy1,0)'
    )
    with ZipFile(output) as package:
        assert package.read("xl/tables/table1.xml") == original_table
        source_root = parse_xml(package.read(_SHEET))
        original_root = parse_xml(original_sheet)
        for root in (source_root, original_root):
            for cell in root.findall(f"{{{_S}}}sheetData/{{{_S}}}row/{{{_S}}}c"):
                if cell.find(f"{{{_S}}}f") is not None:
                    cached = cell.find(f"{{{_S}}}v")
                    if cached is not None:
                        cell.remove(cached)
        assert serialize_xml(source_root) == serialize_xml(original_root)
