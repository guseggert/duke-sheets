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
