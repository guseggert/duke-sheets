//! Tests for reading row heights and column widths from XLSX files created by Excel.

use crate::{cleanup_fixture, ensure_vm_temp_dir, excel_bridge, pull_file_from_vm, temp_fixture};
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_custom_row_height() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Tall row").expect("set value");
        wb.set_row_height(0, 30.0).expect("set row height");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let height = sheet.row_height(0);
    assert!(
        (height - 30.0).abs() < 1.5,
        "Row height should be ~30, got {}",
        height
    );

    cleanup_fixture(&fixture);
}

#[test]
fn test_custom_column_width() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Wide column").expect("set value");
        wb.set_column_width(0, 20.0).expect("set column width");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let width = sheet.column_width(0);
    // Excel's column width conversion is approximate — allow some tolerance
    assert!(
        width > 15.0,
        "Column width should be significantly wider than default (8.43), got {}",
        width
    );

    cleanup_fixture(&fixture);
}

#[test]
fn test_multiple_row_heights() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Row 1").expect("set value");
        wb.set_cell_value("A2", "Row 2").expect("set value");
        wb.set_cell_value("A3", "Row 3").expect("set value");
        wb.set_row_height(0, 25.0).expect("set row 1 height");
        wb.set_row_height(2, 40.0).expect("set row 3 height");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let custom = sheet.custom_row_heights();
    assert!(
        custom.len() >= 2,
        "Should have at least 2 custom row heights, got {}",
        custom.len()
    );

    cleanup_fixture(&fixture);
}

#[test]
fn test_hidden_row() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Visible").expect("set value");
        wb.set_cell_value("A2", "Hidden").expect("set value");
        wb.set_cell_value("A3", "Visible").expect("set value");
        wb.set_row_hidden(1, true).expect("hide row 2");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    assert!(!sheet.is_row_hidden(0), "Row 0 should not be hidden");
    assert!(sheet.is_row_hidden(1), "Row 1 should be hidden");
    assert!(!sheet.is_row_hidden(2), "Row 2 should not be hidden");

    cleanup_fixture(&fixture);
}
