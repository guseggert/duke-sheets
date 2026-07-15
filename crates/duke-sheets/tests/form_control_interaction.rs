use duke_sheets::{
    CellMarker, CellValue, CheckState, ChildTransform, DrawingAnchor, DrawingKind, DrawingMeta,
    DrawingObject, FormControl, FormControlKind, Group, GroupChild, GroupTransform, Workbook,
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

fn checkbox(caption: &str, state: CheckState, link: Option<&str>) -> DrawingObject {
    DrawingObject::form_control(FormControl::new(FormControlKind::Checkbox {
        caption: caption.into(),
        state,
        cell_link: link.map(str::to_string),
        no_3d: false,
    }))
}

fn check_states(workbook: &Workbook, sheet: usize) -> Vec<CheckState> {
    workbook
        .worksheet(sheet)
        .unwrap()
        .placed_form_controls()
        .iter()
        .filter_map(|placed| match placed.control.kind {
            FormControlKind::Checkbox { state, .. }
            | FormControlKind::OptionButton { state, .. } => Some(state),
            _ => None,
        })
        .collect()
}

// features: Form control: option (radio) button
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
    ).unwrap();
    sheet.add_drawing(radio("One", true, Some("$F$2")).with_anchor(anchor(1, 1, 3, 2))).unwrap();
    sheet.add_drawing(radio("Two", false, None).with_anchor(anchor(1, 3, 3, 4))).unwrap();

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

/// Interacting with one control must not rewrite the linked cells of
/// unrelated controls, matching Excel: a stale disagreeing link
/// elsewhere stays untouched until a full sync or save.
#[test]
fn interaction_sync_is_scoped_to_the_affected_controls() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.add_drawing(
        DrawingObject::form_control(FormControl::new(FormControlKind::Checkbox {
            caption: "target".into(),
            state: CheckState::Unchecked,
            cell_link: Some("$A$1".into()),
            no_3d: false,
        }))
        .with_anchor(anchor(0, 0, 2, 1)),
    ).unwrap();
    sheet.add_drawing(
        DrawingObject::form_control(FormControl::new(FormControlKind::Checkbox {
            caption: "unrelated".into(),
            state: CheckState::Unchecked,
            cell_link: Some("$B$1".into()),
            no_3d: false,
        }))
        .with_anchor(anchor(0, 2, 2, 3)),
    ).unwrap();
    // B1 disagrees with the unrelated unchecked checkbox; Excel would
    // only reconcile it on load/save, never on another control's click.
    sheet.set_cell_value("B1", true).unwrap();

    let result = workbook
        .set_form_control_check_state(0, &[0], CheckState::Checked)
        .unwrap();
    assert_eq!(result.controls_changed, 1);
    assert_eq!(result.linked_cells_changed, 1);
    let sheet = workbook.worksheet(0).unwrap();
    assert_eq!(
        sheet.get_value("A1").unwrap(),
        duke_sheets::CellValue::Boolean(true),
        "target link updates"
    );
    assert_eq!(
        sheet.get_value("B1").unwrap(),
        duke_sheets::CellValue::Boolean(true),
        "unrelated stale link must stay untouched"
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
    ).unwrap();
    sheet.add_drawing(
        DrawingObject::form_control(FormControl::new(FormControlKind::Button {
            caption: "Run".into(),
        }))
        .with_anchor(anchor(0, 2, 2, 3)),
    ).unwrap();

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

/// Controls sharing the target's linked cell must follow the cell's
/// new value at interaction time. Otherwise the save-time full sync
/// projects the stale peer into the same cell (last in list wins) and
/// silently reverts the user's change.
#[test]
fn toggling_a_shared_link_checkbox_reconciles_its_peers() {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.add_drawing(
        checkbox("X", CheckState::Checked, Some("$A$1")).with_anchor(anchor(0, 0, 2, 1)),
    ).unwrap();
    sheet.add_drawing(
        checkbox("Y", CheckState::Checked, Some("$A$1")).with_anchor(anchor(0, 2, 2, 3)),
    ).unwrap();
    sheet.set_cell_value("A1", true).unwrap();

    let result = workbook
        .set_form_control_check_state(0, &[0], CheckState::Unchecked)
        .unwrap();
    assert_eq!(result.controls_changed, 2, "X toggled, Y reconciled");
    assert_eq!(result.linked_cells_changed, 1);
    assert_eq!(
        workbook.worksheet(0).unwrap().get_value("A1").unwrap(),
        CellValue::Boolean(false)
    );
    assert_eq!(
        check_states(&workbook, 0),
        [CheckState::Unchecked, CheckState::Unchecked],
        "Y must follow the shared linked cell"
    );

    let snapshot = workbook
        .synchronized_for_save()
        .expect("linked controls produce a snapshot");
    assert_eq!(
        snapshot.worksheet(0).unwrap().get_value("A1").unwrap(),
        CellValue::Boolean(false),
        "the full save-time sync must not revert the interaction"
    );
}

/// Same reconciliation across sheets: a peer on another sheet linked
/// to the target's cell (and the target linking cross-sheet) follows
/// the new cell value, mirroring how the full sync resolves links.
#[test]
fn shared_link_reconciliation_follows_cross_sheet_links() {
    let mut workbook = Workbook::new();
    workbook.add_worksheet_with_name("Sheet2").unwrap();
    let sheet1 = workbook.worksheet_mut(0).unwrap();
    sheet1.add_drawing(
        checkbox("remote", CheckState::Checked, Some("Sheet2!$A$1"))
            .with_anchor(anchor(0, 0, 2, 1)),
    ).unwrap();
    let sheet2 = workbook.worksheet_mut(1).unwrap();
    sheet2.add_drawing(
        checkbox("local", CheckState::Checked, Some("$A$1")).with_anchor(anchor(0, 0, 2, 1)),
    ).unwrap();
    sheet2.set_cell_value("A1", true).unwrap();

    let result = workbook
        .set_form_control_check_state(0, &[0], CheckState::Unchecked)
        .unwrap();
    assert_eq!(result.controls_changed, 2);
    assert_eq!(result.linked_cells_changed, 1);
    assert_eq!(
        workbook.worksheet(1).unwrap().get_value("A1").unwrap(),
        CellValue::Boolean(false)
    );
    assert_eq!(
        check_states(&workbook, 1),
        [CheckState::Unchecked],
        "the Sheet2 peer must follow the shared cell"
    );

    let snapshot = workbook.synchronized_for_save().unwrap();
    assert_eq!(
        snapshot.worksheet(1).unwrap().get_value("A1").unwrap(),
        CellValue::Boolean(false)
    );
}

// features: Form control: option (radio) button
#[test]
fn checking_nested_radio_by_path_updates_group_and_linked_cell() {
    fn nested_radio(caption: &str, checked: bool, link: Option<&str>, y_emu: i64) -> GroupChild {
        GroupChild {
            meta: DrawingMeta::default(),
            transform: ChildTransform {
                x_emu: 0,
                y_emu,
                cx_emu: 1000,
                cy_emu: 400,
                ..ChildTransform::default()
            },
            kind: DrawingKind::FormControl(FormControl::new(FormControlKind::OptionButton {
                caption: caption.into(),
                state: if checked {
                    CheckState::Checked
                } else {
                    CheckState::Unchecked
                },
                cell_link: link.map(str::to_string),
                first_in_group: false,
                no_3d: false,
            })),
        }
    }

    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.add_drawing(
        DrawingObject::group(Group {
            transform: GroupTransform {
                child_x_emu: 0,
                child_y_emu: 0,
                child_cx_emu: 1000,
                child_cy_emu: 1000,
                ..GroupTransform::default()
            },
            children: vec![
                nested_radio("One", true, Some("$E$1"), 0),
                nested_radio("Two", false, None, 500),
            ],
        })
        .with_anchor(anchor(0, 0, 4, 8)),
    ).unwrap();

    let result = workbook
        .set_form_control_check_state(0, &[0, 1], CheckState::Checked)
        .unwrap();
    assert_eq!(result.controls_changed, 2);
    assert_eq!(result.linked_cells_changed, 1);

    let controls = workbook.worksheet(0).unwrap().placed_form_controls();
    let states: Vec<_> = controls
        .iter()
        .map(|placed| match placed.control.kind {
            FormControlKind::OptionButton { state, .. } => (placed.path.clone(), state),
            _ => panic!("expected option buttons"),
        })
        .collect();
    assert_eq!(
        states,
        [
            (vec![0, 0], CheckState::Unchecked),
            (vec![0, 1], CheckState::Checked),
        ]
    );
    assert_eq!(
        workbook.worksheet(0).unwrap().get_value("E1").unwrap(),
        CellValue::Number(2.0)
    );
}
