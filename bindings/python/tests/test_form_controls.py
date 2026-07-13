"""Tests for form controls in the unified drawing API."""

import duke_sheets
import pytest


def anchor(from_row=1, from_col=1, to_row=2, to_col=3):
    return duke_sheets.DrawingAnchor(from_row, from_col, to_row, to_col)


def drawing(control, *, name=None, drawing_anchor=None, meta=None):
    if meta is None:
        meta = duke_sheets.DrawingMeta(name=name)
    return duke_sheets.Drawing(
        control,
        anchor=drawing_anchor or anchor(),
        meta=meta,
    )


def all_controls():
    return [
        duke_sheets.FormControl.button("Run"),
        duke_sheets.FormControl.checkbox(
            "Check", state="checked", cell_link="$D$2", no_3d=True
        ),
        duke_sheets.FormControl.option_button("Option"),
        duke_sheets.FormControl.label("Label"),
        duke_sheets.FormControl.group_box("Group"),
        duke_sheets.FormControl.list_box(
            input_range="$A$1:$A$3",
            selection="multi",
            selected=[0, 2],
        ),
        duke_sheets.FormControl.dropdown(
            input_range="$A$1:$A$3", selected=2, lines=8
        ),
        duke_sheets.FormControl.scrollbar(
            value=5, min=0, max=10, increment=1, page=2
        ),
        duke_sheets.FormControl.spinner(value=2, min=0, max=10),
    ]


def test_form_controls_are_flattened_drawing_wrappers():
    workbook = duke_sheets.Workbook()
    sheet = workbook.get_sheet(0)
    for index, control in enumerate(all_controls()):
        sheet.add_drawing(drawing(control, name=f"control {index}"))

    assert sheet.form_control_count == 9
    assert [item.kind for item in sheet.form_controls] == ["form_control"] * 9
    assert [item.form_control.kind for item in sheet.form_controls] == [
        "button",
        "checkbox",
        "option_button",
        "label",
        "group_box",
        "list_box",
        "dropdown",
        "scrollbar",
        "spinner",
    ]
    assert sheet.form_controls[1].drawing_path == [1]
    assert sheet.form_controls[1].form_control.state == "checked"
    assert sheet.form_controls[5].form_control.selected == [0, 2]

    replacement = drawing(duke_sheets.FormControl.label("Replaced"))
    sheet.set_drawing([0], replacement)
    assert sheet.form_controls[0].form_control.caption_text == "Replaced"
    sheet.remove_drawing([0])
    assert sheet.form_control_count == 8
    assert sheet.form_controls[0].form_control.kind == "checkbox"
    with pytest.raises(IndexError):
        sheet.remove_drawing([99])

    assert not hasattr(sheet, "add_form_control")
    assert not hasattr(sheet, "set_form_control")
    assert not hasattr(sheet, "remove_form_control")
    assert callable(workbook.sync_form_controls)
    assert callable(workbook.sync_form_controls_from_linked_cells)


@pytest.mark.parametrize("extension", ["xlsx", "xlsb", "xls"])
def test_form_controls_round_trip(tmp_path, extension):
    workbook = duke_sheets.Workbook()
    sheet = workbook.get_sheet(0)
    for control in all_controls():
        sheet.add_drawing(drawing(control))

    path = str(tmp_path / f"controls.{extension}")
    workbook.save(path)
    reopened = duke_sheets.Workbook.open(path).get_sheet(0)
    assert reopened.form_control_count == 9

    checkbox = reopened.form_controls[1].form_control
    assert (checkbox.kind, checkbox.caption_text, checkbox.state) == (
        "checkbox",
        "Check",
        "checked",
    )
    assert checkbox.no_3d is True
    assert checkbox.cell_link == "$D$2"

    option = reopened.form_controls[2].form_control
    assert (option.state, option.first_in_group) == ("unchecked", True)

    list_box = reopened.form_controls[5].form_control
    assert (list_box.input_range, list_box.selection, list_box.selected) == (
        "$A$1:$A$3",
        "multi",
        [0, 2],
    )

    dropdown = reopened.form_controls[6].form_control
    assert (dropdown.input_range, dropdown.selected, dropdown.lines) == (
        "$A$1:$A$3",
        [2],
        8,
    )
    assert reopened.get_cell("D2").as_boolean() is True


def test_radio_interaction_updates_siblings_and_linked_cell():
    workbook = duke_sheets.Workbook()
    sheet = workbook.get_sheet(0)
    sheet.add_drawing(
        drawing(
            duke_sheets.FormControl.group_box("Choose"),
            drawing_anchor=anchor(0, 0, 6, 4),
        )
    )
    sheet.add_drawing(
        drawing(
            duke_sheets.FormControl.option_button(
                "One", state="checked", cell_link="$D$2"
            ),
            drawing_anchor=anchor(1, 1, 2, 2),
        )
    )
    sheet.add_drawing(
        drawing(
            duke_sheets.FormControl.option_button(
                "Two", state="unchecked", cell_link="$D$2"
            ),
            drawing_anchor=anchor(3, 1, 4, 2),
        )
    )

    result = sheet.set_form_control_check_state([2], "checked")
    assert result.controls_changed == 2
    assert result.linked_cells_changed == 1
    assert sheet.get_cell("D2").as_number() == 2
    assert [item.form_control.state for item in sheet.form_controls[1:]] == [
        "unchecked",
        "checked",
    ]


def test_rich_caption_macro_unknown_and_validation():
    caption = duke_sheets.DrawingText(
        [
            duke_sheets.RichTextRun("Run "),
            duke_sheets.RichTextRun("now", font=duke_sheets.RunFont(bold=True)),
        ],
        horizontal_alignment="center",
    )
    control = duke_sheets.FormControl.button(caption, macro_name="RunReport")
    assert control.caption.plain_text == "Run now"
    assert control.caption.runs[1].font.bold is True
    assert control.macro_name == "RunReport"

    unknown = duke_sheets.FormControl.unknown("EditBox", "Value")
    assert unknown.kind == "unknown"
    assert unknown.object_type == "EditBox"

    with pytest.raises(ValueError, match="mixed"):
        duke_sheets.FormControl.option_button("Bad", state="mixed")
    with pytest.raises(ValueError, match="sorted and unique"):
        duke_sheets.FormControl.list_box(selection="multi", selected=[3, 1])
