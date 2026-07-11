use duke_sheets::{
    CellError, CellValue, CheckState, FormControl, FormControlKind, ListSelection, Workbook,
    WorkbookExt, WorkbookOpenOptions, WorkbookSaveOptions,
};

fn linked_controls_workbook() -> Workbook {
    let mut workbook = Workbook::new();
    let sheet = workbook.worksheet_mut(0).unwrap();
    sheet.set_cell_formula("A8", "=FALSE").unwrap();
    sheet.set_cell_value("A3", 99.0).unwrap();
    for (row, item) in ["Alpha", "Beta", "Gamma", "Delta"].iter().enumerate() {
        sheet.set_cell_value_at(row as u32, 7, *item).unwrap();
    }
    let kinds = [
        FormControlKind::Checkbox {
            caption: "mixed".into(),
            state: CheckState::Mixed,
            cell_link: Some("$A$1".into()),
            no_3d: false,
        },
        FormControlKind::ListBox {
            input_range: Some("$H$1:$H$4".into()),
            cell_link: Some("$A$2".into()),
            selection: ListSelection::Single,
            selected: vec![2],
            no_3d: false,
        },
        FormControlKind::ListBox {
            input_range: Some("$H$1:$H$4".into()),
            cell_link: Some("$A$3".into()),
            selection: ListSelection::Multi,
            selected: vec![0, 2],
            no_3d: false,
        },
        FormControlKind::Dropdown {
            input_range: Some("$H$1:$H$4".into()),
            cell_link: Some("$A$4".into()),
            selected: None,
            lines: 8,
            no_3d: false,
        },
        FormControlKind::Scrollbar {
            value: 55,
            min: 0,
            max: 100,
            increment: 1,
            page: 10,
            horizontal: false,
            cell_link: Some("$A$5".into()),
        },
        FormControlKind::Spinner {
            value: 18,
            min: 0,
            max: 30,
            increment: 1,
            cell_link: Some("$A$6".into()),
        },
        FormControlKind::OptionButton {
            caption: "one".into(),
            state: CheckState::Unchecked,
            cell_link: Some("$A$7".into()),
            first_in_group: false,
            no_3d: false,
        },
        FormControlKind::OptionButton {
            caption: "two".into(),
            state: CheckState::Checked,
            cell_link: Some("$A$7".into()),
            first_in_group: false,
            no_3d: false,
        },
        FormControlKind::Checkbox {
            caption: "formula overwrite".into(),
            state: CheckState::Checked,
            cell_link: Some("$A$8".into()),
            no_3d: false,
        },
    ];
    for kind in kinds {
        sheet.add_form_control(FormControl::new(kind));
    }
    workbook
}

#[test]
fn high_level_save_synchronizes_form_control_linked_cells() {
    for extension in ["xlsx", "xlsb", "xls"] {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join(format!("controls.{extension}"));
        let workbook = linked_controls_workbook();
        workbook.save(&path).unwrap();

        assert!(workbook.worksheet(0).unwrap().has_formula_at(7, 0));

        let reopened = Workbook::open(&path).unwrap();
        let sheet = reopened.worksheet(0).unwrap();
        assert_eq!(
            sheet.get_value("A1").unwrap(),
            CellValue::Error(CellError::Na)
        );
        assert_eq!(sheet.get_value("A2").unwrap(), CellValue::Number(3.0));
        assert_eq!(sheet.get_value("A3").unwrap(), CellValue::Number(0.0));
        assert_eq!(sheet.get_value("A4").unwrap(), CellValue::Empty);
        assert_eq!(sheet.get_value("A5").unwrap(), CellValue::Number(55.0));
        assert_eq!(sheet.get_value("A6").unwrap(), CellValue::Number(18.0));
        assert_eq!(sheet.get_value("A7").unwrap(), CellValue::Number(2.0));
        assert_eq!(sheet.get_value("A8").unwrap(), CellValue::Boolean(true));
    }
}

#[test]
fn csv_save_synchronizes_form_control_linked_cells() {
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("controls.csv");
    let workbook = linked_controls_workbook();
    workbook.save(&path).unwrap();

    let csv = std::fs::read_to_string(&path).unwrap();
    let first_cells: Vec<&str> = csv
        .lines()
        .map(|line| line.split(',').next().unwrap_or(""))
        .collect();
    // Rows 2/5/7 hold the single-select index, scrollbar value, and
    // radio-group index that synchronization writes into column A.
    assert_eq!(first_cells[1], "3");
    assert_eq!(first_cells[4], "55");
    assert_eq!(first_cells[6], "2");
    // The caller's workbook is untouched.
    assert_eq!(
        workbook.worksheet(0).unwrap().get_value("A2").unwrap(),
        CellValue::Empty
    );
}

#[test]
fn encrypted_save_synchronizes_form_control_linked_cells() {
    const PASSWORD: &str = "linked-cells";
    for extension in ["xlsx", "xls"] {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join(format!("controls.{extension}"));
        let workbook = linked_controls_workbook();
        workbook
            .save_with(&path, &WorkbookSaveOptions::new().password(PASSWORD))
            .unwrap();

        let reopened =
            Workbook::open_with(&path, &WorkbookOpenOptions::default().password(PASSWORD)).unwrap();
        let sheet = reopened.worksheet(0).unwrap();
        assert_eq!(sheet.get_value("A2").unwrap(), CellValue::Number(3.0));
        assert_eq!(sheet.get_value("A8").unwrap(), CellValue::Boolean(true));
    }
}
