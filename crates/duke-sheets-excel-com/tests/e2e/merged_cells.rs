//! Tests for reading merged cells from XLSX files created by Excel.

use crate::{cleanup_fixture, ensure_vm_temp_dir, excel_bridge, pull_file_from_vm, temp_fixture};
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_merged_cells_horizontal() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Merged horizontal")
            .expect("set value");
        wb.merge_range("A1:C1").expect("merge range");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let regions = sheet.merged_regions();
    assert_eq!(
        regions.len(),
        1,
        "Should have 1 merged region, got {}",
        regions.len()
    );
    let r = &regions[0];
    assert_eq!(r.start.row, 0, "Start row should be 0");
    assert_eq!(r.start.col, 0, "Start col should be 0 (A)");
    assert_eq!(r.end.row, 0, "End row should be 0");
    assert_eq!(r.end.col, 2, "End col should be 2 (C)");

    cleanup_fixture(&fixture);
}

#[test]
fn test_merged_cells_vertical() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Merged vertical")
            .expect("set value");
        wb.merge_range("A1:A3").expect("merge range");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let regions = sheet.merged_regions();
    assert_eq!(regions.len(), 1, "Should have 1 merged region");
    let r = &regions[0];
    assert_eq!(r.start.row, 0);
    assert_eq!(r.start.col, 0);
    assert_eq!(r.end.row, 2, "End row should be 2 (row 3)");
    assert_eq!(r.end.col, 0);

    cleanup_fixture(&fixture);
}

#[test]
fn test_merged_cells_block() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Merged block").expect("set value");
        wb.merge_range("A1:C3").expect("merge range");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let regions = sheet.merged_regions();
    assert_eq!(regions.len(), 1, "Should have 1 merged region");
    let r = &regions[0];
    assert_eq!(r.start.row, 0);
    assert_eq!(r.start.col, 0);
    assert_eq!(r.end.row, 2);
    assert_eq!(r.end.col, 2);

    cleanup_fixture(&fixture);
}
