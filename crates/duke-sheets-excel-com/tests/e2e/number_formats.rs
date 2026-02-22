//! Tests for reading number formats from XLSX files created by Excel.

use crate::{cleanup_fixture, ensure_vm_temp_dir, excel_bridge, pull_file_from_vm, temp_fixture};
use duke_sheets_core::NumberFormat;
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_percentage_format() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", 0.1234).expect("set value");
        wb.set_number_format("A1", "0.00%")
            .expect("set number format");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    let fmt = style.number_format.format_string();
    assert!(
        fmt.contains('%'),
        "Number format should contain '%', got: {fmt}"
    );

    cleanup_fixture(&fixture);
}

#[test]
fn test_currency_format() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", 1234.56).expect("set value");
        wb.set_number_format("A1", "\"$\"#,##0.00")
            .expect("set number format");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(
        style.number_format != NumberFormat::General,
        "Number format should not be General"
    );

    cleanup_fixture(&fixture);
}

#[test]
fn test_custom_decimal_format() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", 1234.5678).expect("set value");
        wb.set_number_format("A1", "#,##0.00")
            .expect("set number format");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(
        style.number_format != NumberFormat::General,
        "Should not be General"
    );
    let fmt = style.number_format.format_string();
    assert!(
        fmt.contains("#,##0") || fmt.contains("0.00"),
        "Format should be decimal, got: {fmt}"
    );

    cleanup_fixture(&fixture);
}

#[test]
fn test_date_format() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        // Excel serial date for 2024-03-15
        wb.set_cell_value("A1", 45366.0).expect("set value");
        wb.set_number_format("A1", "YYYY-MM-DD")
            .expect("set number format");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(
        style.number_format != NumberFormat::General,
        "Should not be General"
    );

    cleanup_fixture(&fixture);
}
