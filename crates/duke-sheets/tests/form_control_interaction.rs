use duke_sheets::{
    CellMarker, CheckState, DrawingAnchor, DrawingObject, FormControl, FormControlKind, Workbook,
};

fn anchor(from_col: u16, from_row: u32, to_col: u16, to_row: u32) -> DrawingAnchor {
    DrawingAnchor::TwoCell {
        from: CellMarker {
            col: from_col,
            row: from_row,
            ..CellMarker::default()
        },
        to: CellMarker {
            col: to_col,
            row: to_row,
            ..CellMarker::default()
        },
        edit_as: None,
    }
}

fn radio(caption: &str, checked: bool, link: Option<&str>) -> DrawingObject {
    DrawingObject::form_control(FormControl::new(FormControlKind::OptionButton {
        caption: caption.into(),
        state: if checked {
            CheckState::Checked
        } else {
            CheckState::Unchecked
        },
        cell_link: link.map(str::to_string),
        first_in_group: false,
        no_3d: false,
    }))
}

#[test]
fn checking_radio_updates_group_and_linked_cell() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.add_drawing(
        DrawingObject::form_control(FormControl::new(FormControlKind::GroupBox {
            caption: "Choices".into(),
            no_3d: false,
        }))
        .with_anchor(anchor(0, 0, 4, 8)),
    );
    sheet.add_drawing(radio("One", true, Some("$F$2")).with_anchor(anchor(1, 1, 3, 2)));
    sheet.add_drawing(radio("Two", false, None).with_anchor(anchor(1, 3, 3, 4)));

    let result = workbook
        .set_form_control_check_state(0, &[2], CheckState::Checked)
        .unwrap();
    assert_eq!(result.controls_changed, 2);
    assert_eq!(result.linked_cells_changed, 1);

    let controls = workbook.worksheet(0).unwrap().placed_form_controls();
    let states: Vec<_> = controls
        .iter()
        .filter_map(|placed| match placed.control.kind {
            FormControlKind::OptionButton { state, .. } => Some(state),
            _ => None,
        })
        .collect();
    assert_eq!(states, [CheckState::Unchecked, CheckState::Checked]);
    assert_eq!(
        workbook.worksheet(0).unwrap().get_value("F2").unwrap(),
        duke_sheets::CellValue::Number(2.0)
    );
}

#[test]
fn checkbox_and_invalid_targets_use_semantic_validation() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.add_drawing(
        DrawingObject::form_control(FormControl::new(FormControlKind::Checkbox {
            caption: "Enabled".into(),
            state: CheckState::Unchecked,
            cell_link: Some("$A$1".into()),
            no_3d: false,
        }))
        .with_anchor(anchor(0, 0, 2, 1)),
    );
    sheet.add_drawing(
        DrawingObject::form_control(FormControl::new(FormControlKind::Button {
            caption: "Run".into(),
        }))
        .with_anchor(anchor(0, 2, 2, 3)),
    );

    let result = workbook
        .set_form_control_check_state(0, &[0], CheckState::Mixed)
        .unwrap();
    assert_eq!(result.controls_changed, 1);
    assert_eq!(result.linked_cells_changed, 1);
    assert_eq!(
        workbook.worksheet(0).unwrap().get_value("A1").unwrap(),
        duke_sheets::CellValue::Error(duke_sheets::CellError::Na)
    );

    let error = workbook
        .set_form_control_check_state(0, &[1], CheckState::Checked)
        .unwrap_err();
    assert!(error.to_string().contains("only valid"));
}
