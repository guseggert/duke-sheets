//! Tests for reading data types from XLSX files created by Excel.
//!
//! Each test creates its fixture via Excel COM, saves it, pulls it from
//! the VM, and reads it back with `XlsxReader` to verify.

use crate::{
    assert_has_formula, assert_is_error, assert_number, assert_string, assert_string_contains,
    cleanup_fixture, ensure_vm_temp_dir, excel_bridge, pull_file_from_vm, temp_fixture,
};
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_number_values() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", 42.0).expect("set A1");
        wb.set_cell_value("A2", 3.14159).expect("set A2");
        wb.set_cell_value("A3", -100.0).expect("set A3");
        wb.set_cell_value("A4", 0.0).expect("set A4");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read workbook");
    let sheet = workbook.worksheet(0).expect("worksheet");

    assert_number(&sheet, 0, 0, 42.0, "A1");
    assert_number(&sheet, 1, 0, 3.14159, "A2");
    assert_number(&sheet, 2, 0, -100.0, "A3");
    assert_number(&sheet, 3, 0, 0.0, "A4");

    cleanup_fixture(&fixture);
}

#[test]
fn test_string_values() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Hello").expect("set A1");
        wb.set_cell_value("A2", "World with spaces")
            .expect("set A2");
        // Japanese text: 日本語
        wb.set_cell_value("A3", "Unicode: \u{65e5}\u{672c}\u{8a9e}")
            .expect("set A3");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read workbook");
    let sheet = workbook.worksheet(0).expect("worksheet");

    assert_string(&sheet, 0, 0, "Hello", "A1");
    assert_string(&sheet, 1, 0, "World with spaces", "A2");
    assert_string_contains(&sheet, 2, 0, "\u{65e5}\u{672c}\u{8a9e}", "A3");

    cleanup_fixture(&fixture);
}

#[test]
fn test_boolean_values() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        // Excel preserves booleans directly (unlike LO which sometimes converts to numbers)
        wb.set_cell_value("A1", true).expect("set A1");
        wb.set_cell_value("A2", false).expect("set A2");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read workbook");
    let sheet = workbook.worksheet(0).expect("worksheet");

    // Check effective values - could be Boolean or Number (1/0)
    let mut found_true = false;
    let mut found_false = false;
    for row in 0..5 {
        for col in 0..5 {
            if let Some(cell) = sheet.cell_at(row, col) {
                match cell.value.effective_value() {
                    duke_sheets_core::CellValue::Boolean(b) => {
                        if *b {
                            found_true = true;
                        } else {
                            found_false = true;
                        }
                    }
                    duke_sheets_core::CellValue::Number(n) => {
                        if (*n - 1.0).abs() < 0.001 {
                            found_true = true;
                        } else if n.abs() < 0.001 {
                            found_false = true;
                        }
                    }
                    _ => {}
                }
            }
        }
    }

    assert!(found_true, "Should find TRUE value");
    assert!(found_false, "Should find FALSE value");

    cleanup_fixture(&fixture);
}

#[test]
fn test_formula_values() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", 10.0).expect("set A1");
        wb.set_cell_value("A2", 20.0).expect("set A2");
        wb.set_cell_formula("A3", "=A1+A2").expect("set A3");
        wb.set_cell_formula("A4", "=SUM(A1:A2)").expect("set A4");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read workbook");
    let sheet = workbook.worksheet(0).expect("worksheet");

    assert_has_formula(&sheet, 2, 0, "A3");
    assert_has_formula(&sheet, 3, 0, "A4");

    // Also verify cached values are correct
    assert_number(&sheet, 2, 0, 30.0, "A3 value");
    assert_number(&sheet, 3, 0, 30.0, "A4 value");

    cleanup_fixture(&fixture);
}

#[test]
fn test_error_values() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_formula("A1", "=1/0").expect("set A1"); // #DIV/0!
        wb.set_cell_formula("A2", "=VALUE(\"x\")").expect("set A2"); // #VALUE!

        excel.recalculate().expect("recalculate");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read workbook");
    let sheet = workbook.worksheet(0).expect("worksheet");

    assert_is_error(&sheet, 0, 0, "A1");
    assert_is_error(&sheet, 1, 0, "A2");

    cleanup_fixture(&fixture);
}
