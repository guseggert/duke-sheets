use crate::{cleanup_fixture, ensure_vm_temp_dir, excel_bridge, pull_file_from_vm, temp_fixture};
use duke_sheets_core::CellValue;
use duke_sheets_xlsx::XlsxReader;

#[test]
#[ignore = "spill ghost cells read from Excel output are not tagged SpillTarget; FEATURES.md 'Dynamic array formulas (spill)' is W● pending Excel parity"]
fn test_xlsx_dynamic_array_sequence() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        // =SEQUENCE(3) spills to A1:A3 with values 1,2,3
        wb.set_cell_formula2("A1", "=SEQUENCE(3)")
            .expect("set formula2");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read workbook");
    let sheet = workbook.worksheet(0).expect("worksheet");

    // Anchor cell: Formula with array_result
    let a1 = sheet.get_value("A1").unwrap();
    assert!(
        sheet.get_formula_at(0, 0).is_some(),
        "A1 should be a formula, got {:?}",
        a1
    );
    assert!(
        sheet
            .formula_data_at(0, 0)
            .map(|formula| formula.is_array_formula())
            .unwrap_or(false),
        "A1 should have array_result"
    );
    assert_eq!(a1.as_number(), Some(1.0));

    // Ghost cells: SpillTarget
    let a2 = sheet.get_value("A2").unwrap();
    assert!(
        a2.is_spill_target(),
        "A2 should be SpillTarget, got {:?}",
        a2
    );
    let a3 = sheet.get_value("A3").unwrap();
    assert!(
        a3.is_spill_target(),
        "A3 should be SpillTarget, got {:?}",
        a3
    );

    // Resolved values
    assert_eq!(sheet.get_value_at(1, 0).as_number(), Some(2.0));
    assert_eq!(sheet.get_value_at(2, 0).as_number(), Some(3.0));

    cleanup_fixture(&fixture);
}

#[test]
#[ignore = "spill ghost cells read from Excel output are not tagged SpillTarget; FEATURES.md 'Dynamic array formulas (spill)' is W● pending Excel parity"]
fn test_xlsx_dynamic_array_2d() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        // =SEQUENCE(2,3) spills to A1:C2 with values 1..6
        wb.set_cell_formula2("A1", "=SEQUENCE(2,3)")
            .expect("set formula2");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read workbook");
    let sheet = workbook.worksheet(0).expect("worksheet");

    // Anchor
    assert!(
        sheet
            .formula_data_at(0, 0)
            .map(|formula| formula.is_array_formula())
            .unwrap_or(false),
        "A1 should have array_result"
    );

    // All non-anchor cells are SpillTarget
    for cell_ref in ["B1", "C1", "A2", "B2", "C2"] {
        let val = sheet.get_value(cell_ref).unwrap();
        assert!(
            val.is_spill_target(),
            "{cell_ref} should be SpillTarget, got {:?}",
            val
        );
    }

    // Resolved values: [[1,2,3],[4,5,6]]
    assert_eq!(sheet.get_value_at(0, 0).as_number(), Some(1.0));
    assert_eq!(sheet.get_value_at(0, 1).as_number(), Some(2.0));
    assert_eq!(sheet.get_value_at(0, 2).as_number(), Some(3.0));
    assert_eq!(sheet.get_value_at(1, 0).as_number(), Some(4.0));
    assert_eq!(sheet.get_value_at(1, 1).as_number(), Some(5.0));
    assert_eq!(sheet.get_value_at(1, 2).as_number(), Some(6.0));

    cleanup_fixture(&fixture);
}

#[test]
#[ignore = "spill ghost cells read from Excel output are not tagged SpillTarget; FEATURES.md 'Dynamic array formulas (spill)' is W● pending Excel parity"]
fn test_xlsx_dynamic_array_unique_strings() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "apple").expect("set A1");
        wb.set_cell_value("A2", "banana").expect("set A2");
        wb.set_cell_value("A3", "apple").expect("set A3");
        wb.set_cell_value("A4", "cherry").expect("set A4");
        wb.set_cell_formula2("B1", "=UNIQUE(A1:A4)")
            .expect("set formula2");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read workbook");
    let sheet = workbook.worksheet(0).expect("worksheet");

    // Anchor with formula
    assert!(
        sheet.get_formula_at(0, 1).is_some(),
        "B1 should have formula"
    );
    assert!(
        sheet
            .formula_data_at(0, 1)
            .map(|formula| formula.is_array_formula())
            .unwrap_or(false),
        "B1 should have array_result"
    );

    // Ghost cells are SpillTarget
    assert!(sheet.get_value("B2").unwrap().is_spill_target());
    assert!(sheet.get_value("B3").unwrap().is_spill_target());

    // Resolved string values: apple, banana, cherry (3 unique)
    assert_eq!(
        sheet.get_value_at(0, 1),
        CellValue::String("apple".to_string().into())
    );
    assert_eq!(
        sheet.get_value_at(1, 1),
        CellValue::String("banana".to_string().into())
    );
    assert_eq!(
        sheet.get_value_at(2, 1),
        CellValue::String("cherry".to_string().into())
    );

    cleanup_fixture(&fixture);
}
