"""Tests for form-control bindings."""

import duke_sheets
import pytest


def anchor():
    return duke_sheets.DrawingAnchor(1, 1, 2, 3)


def all_controls():
    a = anchor()
    return [
        duke_sheets.FormControl.button("Run", a),
        duke_sheets.FormControl.checkbox("Check", a, state="checked", no_3d=True),
        duke_sheets.FormControl.option_button("Option", a),
        duke_sheets.FormControl.label("Label", a),
        duke_sheets.FormControl.group_box("Group", a),
        duke_sheets.FormControl.list_box(
            a,
            input_range="$A$1:$A$3",
            selection="multi",
            # Zero-based: first and third items.
            selected=[0, 2],
        ),
        duke_sheets.FormControl.dropdown(
            a, input_range="$A$1:$A$3", selected=2, lines=8
        ),
        duke_sheets.FormControl.scrollbar(
            a, value=5, min=0, max=10, increment=1, page=2
        ),
        duke_sheets.FormControl.spinner(a, value=2, min=0, max=10),
    ]


def test_add_set_remove_form_controls():
    workbook = duke_sheets.Workbook()
    sheet = workbook.get_sheet(0)
    for control in all_controls():
        sheet.add_form_control(control)

    assert sheet.form_control_count == 9
    assert [control.kind for control in sheet.form_controls] == [
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
    assert sheet.form_controls[1].state == "checked"
    assert sheet.form_controls[5].selected == [0, 2]

    sheet.set_form_control(0, duke_sheets.FormControl.label("Replaced", anchor()))
    assert sheet.form_controls[0].caption == "Replaced"
    sheet.remove_form_control(0)
    assert sheet.form_control_count == 8
    assert sheet.form_controls[0].kind == "checkbox"
    with pytest.raises(IndexError):
        sheet.remove_form_control(99)


@pytest.mark.parametrize("extension", ["xlsx", "xlsb", "xls"])
def test_form_controls_round_trip(tmp_path, extension):
    workbook = duke_sheets.Workbook()
    sheet = workbook.get_sheet(0)
    for control in all_controls():
        sheet.add_form_control(control)

    path = str(tmp_path / f"controls.{extension}")
    workbook.save(path)
    reopened = duke_sheets.Workbook.open(path).get_sheet(0)
    assert reopened.form_control_count == 9
    assert reopened.form_controls[1].kind == "checkbox"
    assert reopened.form_controls[5].selected == [0, 2]
    assert reopened.form_controls[6].selected == [2]


def test_invalid_form_control_inputs():
    a = anchor()
    with pytest.raises(ValueError, match="mixed"):
        duke_sheets.FormControl.option_button("Bad", a, state="mixed")
    with pytest.raises(ValueError, match="sorted and unique"):
        duke_sheets.FormControl.list_box(
            a, selection="multi", selected=[3, 1]
        )
