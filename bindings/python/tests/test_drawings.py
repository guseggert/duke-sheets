"""Focused tests for the recursive worksheet drawing API."""

import duke_sheets


def anchor(from_row=0, from_col=0, to_row=2, to_col=2):
    return duke_sheets.DrawingAnchor(from_row, from_col, to_row, to_col)


def child(payload, *, name=None, x=0, y=0):
    return duke_sheets.Drawing(
        payload,
        transform=duke_sheets.ChildTransform(
            x_emu=x,
            y_emu=y,
            cx_emu=100_000,
            cy_emu=50_000,
        ),
        meta=duke_sheets.DrawingMeta(name=name),
    )


def top(payload, *, name=None, meta=None, drawing_anchor=None):
    return duke_sheets.Drawing(
        payload,
        anchor=drawing_anchor or anchor(),
        meta=meta or duke_sheets.DrawingMeta(name=name),
    )


def test_absolute_rect_emu_resolves_group_children():
    workbook = duke_sheets.Workbook()
    sheet = workbook.get_sheet(0)
    sheet.add_drawing(top(duke_sheets.FormControl.label("top"), name="top"))

    group = duke_sheets.DrawingGroup(
        duke_sheets.GroupTransform(child_cx_emu=1000, child_cy_emu=1000),
        [
            duke_sheets.Drawing(
                duke_sheets.FormControl.label("nested"),
                transform=duke_sheets.ChildTransform(
                    x_emu=250, y_emu=500, cx_emu=500, cy_emu=250
                ),
            )
        ],
    )
    sheet.add_drawing(top(group, drawing_anchor=anchor(0, 0, 4, 4)))

    drawings = sheet.drawings
    rect = drawings[0].absolute_rect_emu
    assert (rect.x_emu, rect.y_emu) == (0, 0)
    assert rect.width_emu > 0 and rect.height_emu > 0

    group_rect = drawings[1].absolute_rect_emu
    controls = sheet.form_controls
    assert len(controls) == 2
    nested = controls[1].absolute_rect_emu
    assert nested.x_emu == group_rect.x_emu + group_rect.width_emu // 4
    assert nested.y_emu == group_rect.y_emu + group_rect.height_emu // 2
    assert nested.width_emu == group_rect.width_emu // 2
    assert nested.height_emu == group_rect.height_emu // 4

    # The tree view agrees with the flattened view.
    tree_child = drawings[1].group.children[0].absolute_rect_emu
    assert (tree_child.x_emu, tree_child.width_emu) == (nested.x_emu, nested.width_emu)

    # Drawings constructed in Python have no on-sheet placement yet.
    assert top(duke_sheets.FormControl.label("free")).absolute_rect_emu is None


def test_theme_palette_and_resolve_color():
    workbook = duke_sheets.Workbook()
    palette = workbook.theme_palette
    assert len(palette) == 12
    assert palette[4] == "4F81BD"

    theme = duke_sheets.Color("theme", theme_index=4, tint=0)
    assert workbook.resolve_color(theme) == "4F81BD"
    tinted = duke_sheets.Color("theme", theme_index=4, tint=0.5)
    assert workbook.resolve_color(tinted) == "A7C0DE"
    rgb = duke_sheets.Color("rgb", r=1, g=2, b=3)
    assert workbook.resolve_color(rgb) == "010203"
    assert workbook.resolve_color(duke_sheets.Color("auto")) is None


def test_recursive_paths_and_z_order():
    workbook = duke_sheets.Workbook()
    sheet = workbook.get_sheet(0)
    sheet.add_drawing(top(duke_sheets.Shape("rect"), name="back"))

    nested = child(
        duke_sheets.DrawingGroup(
            duke_sheets.GroupTransform(child_cx_emu=100_000, child_cy_emu=50_000),
            [child(duke_sheets.Shape("ellipse"), name="grandchild")],
        ),
        name="nested group",
    )
    group = duke_sheets.DrawingGroup(
        duke_sheets.GroupTransform(child_cx_emu=200_000, child_cy_emu=100_000),
        [
            child(duke_sheets.FormControl.label("nested"), name="label"),
            nested,
        ],
    )
    sheet.add_drawing(top(group, name="front group"))

    assert [drawing.name for drawing in sheet.drawings] == ["back", "front group"]
    assert sheet.drawings[0].drawing_path == [0]
    assert sheet.drawings[1].drawing_path == [1]
    assert sheet.drawings[1].group.children[0].drawing_path == [1, 0]
    assert sheet.drawings[1].group.children[1].group.children[0].drawing_path == [
        1,
        1,
        0,
    ]
    assert sheet.form_controls[0].drawing_path == [1, 0]


def test_lazy_image_bytes_and_flattened_metadata():
    workbook = duke_sheets.Workbook()
    sheet = workbook.get_sheet(0)
    payload = duke_sheets.EmbeddedImage(
        b"\x89PNG",
        "png",
        width_emu=200_000,
        height_emu=100_000,
    )
    meta = duke_sheets.DrawingMeta(
        name="Logo",
        hidden=True,
        locked=False,
        printable=False,
        alt_text="Company logo",
        title="Brand",
    )
    sheet.add_drawing(top(payload, meta=meta))

    image = sheet.images[0]
    assert image.drawing_path == [0]
    assert image.name == "Logo"
    assert image.hidden is True
    assert image.locked is False
    assert image.printable is False
    assert image.alt_text == "Company logo"
    assert image.title == "Brand"
    assert image.image.format == "png"
    assert not hasattr(image.image, "data")
    assert sheet.drawing_image_data([0]) == b"\x89PNG"
    assert sheet.drawing_svg_data([0]) is None


def test_top_level_and_nested_mutation():
    workbook = duke_sheets.Workbook()
    sheet = workbook.get_sheet(0)
    sheet.add_drawing(top(duke_sheets.Shape("rect"), name="one"))
    sheet.add_drawing(top(duke_sheets.Shape("rect"), name="three"))
    sheet.insert_drawing(1, top(duke_sheets.Shape("ellipse"), name="two"))
    sheet.move_drawing(2, 0)
    assert [drawing.name for drawing in sheet.drawings] == ["three", "one", "two"]

    group = duke_sheets.DrawingGroup(
        duke_sheets.GroupTransform(),
        [child(duke_sheets.Shape("rect"), name="old")],
    )
    sheet.add_drawing(top(group, name="group"))
    sheet.set_drawing(
        [3, 0],
        child(duke_sheets.FormControl.label("replacement"), name="new"),
    )
    assert sheet.drawings[3].group.children[0].kind == "form_control"
    assert sheet.drawings[3].group.children[0].name == "new"
    sheet.remove_drawing([3, 0])
    assert sheet.drawings[3].group.children == []


def test_one_cell_anchor_round_trip(tmp_path):
    workbook = duke_sheets.Workbook()
    sheet = workbook.get_sheet(0)
    one_cell = duke_sheets.DrawingAnchor.one_cell(
        2,
        1,
        width_emu=300_000,
        height_emu=200_000,
        from_row_offset=19_050,
        from_col_offset=9_525,
    )
    sheet.add_drawing(duke_sheets.Drawing(duke_sheets.Shape("rect"), anchor=one_cell))
    first_path = tmp_path / "one_cell.xlsx"
    workbook.save(str(first_path))

    def check(read_anchor):
        assert read_anchor.anchor_type == "one_cell"
        assert read_anchor.from_row == 2
        assert read_anchor.from_col == 1
        assert read_anchor.from_row_offset == 19_050
        assert read_anchor.from_col_offset == 9_525
        assert read_anchor.width_emu == 300_000
        assert read_anchor.height_emu == 200_000
        assert read_anchor.to_row is None
        assert read_anchor.edit_as is None

    reread = duke_sheets.Workbook.open(str(first_path))
    sheet = reread.get_sheet(0)
    drawing = sheet.drawings[0]
    check(drawing.anchor)

    # Identity rewrite must not rewrite the anchor variant or extent.
    sheet.set_drawing([0], drawing)
    second_path = tmp_path / "one_cell_rewritten.xlsx"
    reread.save(str(second_path))
    final = duke_sheets.Workbook.open(str(second_path))
    check(final.get_sheet(0).drawings[0].anchor)


def test_absolute_and_two_cell_anchor_variants():
    workbook = duke_sheets.Workbook()
    sheet = workbook.get_sheet(0)
    sheet.add_drawing(
        duke_sheets.Drawing(
            duke_sheets.Shape("rect"),
            anchor=duke_sheets.DrawingAnchor.absolute(100, 200, 300_000, 400_000),
        )
    )
    sheet.add_drawing(top(duke_sheets.Shape("rect"), name="plain"))

    absolute = sheet.drawings[0].anchor
    assert absolute.anchor_type == "absolute"
    assert (absolute.x_emu, absolute.y_emu) == (100, 200)
    assert (absolute.width_emu, absolute.height_emu) == (300_000, 400_000)
    assert absolute.from_col is None
    assert absolute.edit_as is None

    two_cell = sheet.drawings[1].anchor
    assert two_cell.anchor_type == "two_cell"
    assert (two_cell.from_row, two_cell.from_col) == (0, 0)
    assert (two_cell.to_row, two_cell.to_col) == (2, 2)
    assert two_cell.edit_as == "two_cell"
    assert two_cell.width_emu is None


def test_meta_hidden_defaults_by_drawing_kind():
    comment = duke_sheets.Drawing(
        duke_sheets.DrawingComment(0, 0, "note"),
        anchor=anchor(),
        meta=duke_sheets.DrawingMeta(name="c"),
    )
    assert comment.hidden is True

    shape = duke_sheets.Drawing(
        duke_sheets.Shape("rect"),
        anchor=anchor(),
        meta=duke_sheets.DrawingMeta(name="s"),
    )
    assert shape.hidden is False

    visible = duke_sheets.Drawing(
        duke_sheets.DrawingComment(1, 1, "shown"),
        anchor=anchor(),
        meta=duke_sheets.DrawingMeta(hidden=False),
    )
    assert visible.hidden is False


def test_comment_hidden_default_survives_save(tmp_path):
    workbook = duke_sheets.Workbook()
    sheet = workbook.get_sheet(0)
    sheet.add_drawing(
        duke_sheets.Drawing(
            duke_sheets.DrawingComment(0, 0, "note", author="a"),
            anchor=anchor(),
            meta=duke_sheets.DrawingMeta(name="c"),
        )
    )
    file_path = tmp_path / "comment.xlsx"
    workbook.save(str(file_path))
    reread = duke_sheets.Workbook.open(str(file_path))
    drawing = reread.get_sheet(0).drawings[0]
    assert drawing.kind == "comment"
    assert drawing.hidden is True
    assert drawing.meta.hidden is True


def test_unknown_control_raw_getters():
    unknown = duke_sheets.FormControl.unknown("EditBox", "legacy")
    assert unknown.raw_properties == []
    assert unknown.raw_client_data == []
    assert unknown.raw_obj is None

    checkbox = duke_sheets.FormControl.checkbox("check")
    assert checkbox.raw_properties == []
    # raw_client_data is carried on every control kind.
    assert checkbox.raw_client_data == []
    assert checkbox.raw_obj is None


def test_modeled_control_raw_client_data_survives_reload(tmp_path):
    import zipfile

    workbook = duke_sheets.Workbook()
    sheet = workbook.get_sheet(0)
    sheet.add_drawing(top(duke_sheets.FormControl.checkbox("Audit"), name="cb"))
    file_path = tmp_path / "modeled.xlsx"
    workbook.save(str(file_path))

    # Splice an unmodeled ClientData child into the checkbox shape.
    spliced_path = tmp_path / "spliced.xlsx"
    with zipfile.ZipFile(file_path) as source, zipfile.ZipFile(
        spliced_path, "w"
    ) as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if "vmlDrawing" in item.filename:
                xml = data.decode()
                at = xml.index("ObjectType=\"Checkbox\"")
                close = xml.index("</x:ClientData>", at)
                xml = xml[:close] + "   <x:Disabled/>\n  " + xml[close:]
                data = xml.encode()
            target.writestr(item, data)

    reopened = duke_sheets.Workbook.open(str(spliced_path))
    control = reopened.get_sheet(0).form_controls[0].form_control
    assert control.kind == "checkbox"
    assert control.raw_client_data == [b"<x:Disabled/>"]

    # The passthrough survives a save from the Python model.
    resaved_path = tmp_path / "resaved.xlsx"
    reopened.save(str(resaved_path))
    resaved = duke_sheets.Workbook.open(str(resaved_path))
    control = resaved.get_sheet(0).form_controls[0].form_control
    assert control.raw_client_data == [b"<x:Disabled/>"]


def test_unknown_control_raw_getters_return_bytes_after_reload(tmp_path):
    workbook = duke_sheets.Workbook()
    sheet = workbook.get_sheet(0)
    sheet.add_drawing(
        top(duke_sheets.FormControl.unknown("EditBox", "legacy"), name="u")
    )
    file_path = tmp_path / "unknown.xlsx"
    workbook.save(str(file_path))
    control = (
        duke_sheets.Workbook.open(str(file_path)).get_sheet(0).form_controls[0].form_control
    )
    assert control.kind == "unknown"
    assert isinstance(control.raw_properties, list)
    assert all(
        isinstance(name, str) and isinstance(value, str)
        for name, value in control.raw_properties
    )
    assert isinstance(control.raw_client_data, list)
    assert all(isinstance(fragment, bytes) for fragment in control.raw_client_data)
    assert control.raw_obj is None or isinstance(control.raw_obj, bytes)


def test_chart_getters_return_drawing_wrappers():
    workbook = duke_sheets.Workbook()
    sheet = workbook.get_sheet(0)
    sheet.add_drawing(top(duke_sheets.Chart("line", title="Trend"), name="chart"))
    sheet.add_drawing(
        top(duke_sheets.ChartEx("waterfall", title="Bridge"), name="chart ex")
    )

    assert sheet.charts[0].kind == "chart"
    assert sheet.charts[0].drawing_path == [0]
    assert sheet.charts[0].chart.title == "Trend"
    assert sheet.charts_ex[0].kind == "chart_ex"
    assert sheet.charts_ex[0].drawing_path == [1]
    assert sheet.charts_ex[0].chart_ex.layout == "waterfall"


def _reload(workbook, path):
    workbook.save(str(path))
    return duke_sheets.Workbook.open(str(path))


def test_chart_ex_style_part_is_modelled_not_replayed(tmp_path):
    """A written chartEx always carries the style pair Excel demands."""
    workbook = duke_sheets.Workbook()
    workbook.get_sheet(0).add_drawing(top(duke_sheets.ChartEx("waterfall")))

    chart_ex = _reload(workbook, tmp_path / "cx.xlsx").get_sheet(0).charts_ex[0].chart_ex
    style = chart_ex.style
    assert style is not None, "a written chartEx must carry a style part"
    assert style.raw is None
    assert isinstance(style.id, int)
    # The 29 required CT_StyleEntry of MS-ODRAWXML 5.15 plus the
    # optional dataLabelCallout, which Excel's own part also carries.
    assert len(style.entries) == 30
    assert "dataLabelCallout" in style.entries
    assert chart_ex.color_style is not None
    assert chart_ex.color_style.raw is None

    entry = style.entries["dataPoint"]
    assert isinstance(entry.line_reference.idx, int)
    assert isinstance(entry.fill_reference.idx, int)
    assert isinstance(entry.effect_reference.idx, int)
    assert entry.font_collection in ("major", "minor", "none")


def test_plain_chart_style_is_unset_when_authored(tmp_path):
    """Excel needs no style sibling for a plain chart, so none is emitted."""
    workbook = duke_sheets.Workbook()
    workbook.get_sheet(0).add_drawing(top(duke_sheets.Chart("line", title="Trend")))

    chart = _reload(workbook, tmp_path / "plain.xlsx").get_sheet(0).charts[0].chart
    assert chart.style is None
    assert chart.color_style is None


def test_plain_chart_style_part_is_read_when_the_file_has_one(tmp_path):
    """A plain chart's style is surfaced too, not just a chartEx's.

    The authoring surface cannot set a style, so the fixture is built by
    giving a written package the style sibling Excel would have written
    for a plain chart.
    """
    import re
    import shutil
    import zipfile

    workbook = duke_sheets.Workbook()
    workbook.get_sheet(0).add_drawing(top(duke_sheets.ChartEx("waterfall")))
    workbook.get_sheet(0).add_drawing(top(duke_sheets.Chart("line", title="Trend")))
    source = tmp_path / "source.xlsx"
    workbook.save(str(source))

    # Borrow the style the writer emitted for the chartEx, and hand it to
    # the plain chart under the name and relationship a plain chart uses.
    with zipfile.ZipFile(source) as zf:
        names = zf.namelist()
        style_name = next(n for n in names if re.fullmatch(r"xl/charts/style\d+\.xml", n))
        style_xml = zf.read(style_name)
        chart_name = next(n for n in names if re.fullmatch(r"xl/charts/chart\d+\.xml", n))
        entries = {n: zf.read(n) for n in names}

    chart_num = re.fullmatch(r"xl/charts/chart(\d+)\.xml", chart_name).group(1)
    new_style = f"xl/charts/style{chart_num}.xml"
    assert new_style not in entries, "the plain chart's style slot must be free"
    entries[new_style] = style_xml
    rels_name = f"xl/charts/_rels/chart{chart_num}.xml.rels"
    rel = (
        f'<Relationship Id="rId9000" Type="http://schemas.microsoft.com/'
        f'office/2011/relationships/chartStyle" Target="style{chart_num}.xml"/>'
    )
    existing = entries.get(rels_name)
    if existing is None:
        entries[rels_name] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
            f'2006/relationships">{rel}</Relationships>'
        ).encode()
    else:
        entries[rels_name] = existing.replace(b"</Relationships>", rel.encode() + b"</Relationships>")
    ct = entries["[Content_Types].xml"].replace(
        b"</Types>",
        f'<Override PartName="/{new_style}" ContentType="application/vnd.ms-office.chartstyle+xml"/>'.encode()
        + b"</Types>",
    )
    entries["[Content_Types].xml"] = ct

    patched = tmp_path / "patched.xlsx"
    with zipfile.ZipFile(patched, "w", zipfile.ZIP_DEFLATED) as out:
        for name, data in entries.items():
            out.writestr(name, data)

    chart = duke_sheets.Workbook.open(str(patched)).get_sheet(0).charts[0].chart
    assert chart.style is not None, "the plain chart's style part must be read"
    assert chart.style.raw is None, "it must be modelled, not replayed raw"
    assert len(chart.style.entries) == 30
    assert chart.style.entries["chartArea"].font_collection in ("major", "minor", "none")

    # And it survives being written back out.
    final = tmp_path / "final.xlsx"
    reread = duke_sheets.Workbook.open(str(patched))
    reread.save(str(final))
    again = duke_sheets.Workbook.open(str(final)).get_sheet(0).charts[0].chart
    assert again.style is not None, "the style must survive a rewrite"
    assert len(again.style.entries) == 30
    shutil.rmtree(tmp_path, ignore_errors=True)
