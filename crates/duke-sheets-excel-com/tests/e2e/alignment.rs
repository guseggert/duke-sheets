//! Tests for reading alignment properties from XLSX files created by Excel.

use crate::{cleanup_fixture, ensure_vm_temp_dir, excel_bridge, pull_file_from_vm, temp_fixture};
use duke_sheets_core::{HorizontalAlignment, VerticalAlignment};
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_horizontal_center() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Centered").expect("set value");
        // xlCenter = -4108
        wb.set_horizontal_alignment("A1", -4108)
            .expect("set alignment");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(style.alignment.horizontal, HorizontalAlignment::Center);

    cleanup_fixture(&fixture);
}

#[test]
fn test_horizontal_right() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Right").expect("set value");
        // xlRight = -4152
        wb.set_horizontal_alignment("A1", -4152)
            .expect("set alignment");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(style.alignment.horizontal, HorizontalAlignment::Right);

    cleanup_fixture(&fixture);
}

#[test]
fn test_vertical_bottom() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Bottom").expect("set value");
        wb.set_row_height(1, 40.0).expect("set row height");
        // xlBottom = -4107
        wb.set_vertical_alignment("A1", -4107)
            .expect("set alignment");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(style.alignment.vertical, VerticalAlignment::Bottom);

    cleanup_fixture(&fixture);
}

#[test]
fn test_wrap_text() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "This is a long text that should wrap")
            .expect("set value");
        wb.set_wrap_text("A1", true).expect("set wrap text");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.alignment.wrap_text, "Wrap text should be true");

    cleanup_fixture(&fixture);
}

#[test]
fn test_shrink_to_fit() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Shrink").expect("set value");
        wb.set_shrink_to_fit("A1", true).expect("set shrink to fit");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(
        style.alignment.shrink_to_fit,
        "Shrink to fit should be true"
    );

    cleanup_fixture(&fixture);
}

#[test]
fn test_rotation() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Rotated 45").expect("set value");
        wb.set_row_height(1, 60.0).expect("set row height");
        wb.set_rotation("A1", 45).expect("set rotation");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(
        style.alignment.rotation, 45,
        "Rotation should be 45 degrees"
    );

    cleanup_fixture(&fixture);
}

#[test]
fn test_indent() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Indented").expect("set value");
        wb.set_indent("A1", 2).expect("set indent");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(
        style.alignment.indent >= 1,
        "Indent should be >= 1, got {}",
        style.alignment.indent
    );

    cleanup_fixture(&fixture);
}
