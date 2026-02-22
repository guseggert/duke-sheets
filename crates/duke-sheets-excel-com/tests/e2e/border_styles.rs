//! Tests for reading border styles from XLSX files created by Excel.

use crate::{cleanup_fixture, ensure_vm_temp_dir, excel_bridge, pull_file_from_vm, temp_fixture};
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_thin_border_all_sides() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Thin border").expect("set value");
        // xlContinuous = 1 (line style), xlThin = 2 (weight), black color
        wb.set_border_all("A1", 1, 2, 0x000000).expect("set border");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.border.left.is_some(), "Should have left border");
    assert!(style.border.right.is_some(), "Should have right border");
    assert!(style.border.top.is_some(), "Should have top border");
    assert!(style.border.bottom.is_some(), "Should have bottom border");

    cleanup_fixture(&fixture);
}

#[test]
fn test_border_color() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Red border").expect("set value");
        // xlContinuous = 1, xlMedium = -4138, red color
        wb.set_border_all("A1", 1, -4138, 0xFF0000)
            .expect("set border");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");

    let edge = style
        .border
        .top
        .as_ref()
        .or(style.border.left.as_ref())
        .expect("Should have a border");
    let (r, _, _) = edge.color.to_rgb();
    assert!(r > 200, "Expected red border color");

    cleanup_fixture(&fixture);
}

#[test]
fn test_individual_border_left_only() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Left only").expect("set value");
        // xlEdgeLeft = 7, xlContinuous = 1, xlThin = 2
        wb.set_border_edge("A1", 7, 1, 2, 0x000000)
            .expect("set left border");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.border.left.is_some(), "Should have left border");
    assert!(style.border.right.is_none(), "Should NOT have right border");
    assert!(style.border.top.is_none(), "Should NOT have top border");
    assert!(
        style.border.bottom.is_none(),
        "Should NOT have bottom border"
    );

    cleanup_fixture(&fixture);
}

#[test]
fn test_mixed_border_sides() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Mixed").expect("set value");
        // xlEdgeTop = 8, xlContinuous = 1, xlThin = 2, red
        wb.set_border_edge("A1", 8, 1, 2, 0xFF0000)
            .expect("set top border");
        // xlEdgeBottom = 9, xlContinuous = 1, xlThick = 4, blue
        wb.set_border_edge("A1", 9, 1, 4, 0x0000FF)
            .expect("set bottom border");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.border.top.is_some(), "Should have top border");
    assert!(style.border.bottom.is_some(), "Should have bottom border");

    let top = style.border.top.as_ref().unwrap();
    let (r, _, _) = top.color.to_rgb();
    assert!(r > 200, "Top border should be red");

    let bottom = style.border.bottom.as_ref().unwrap();
    let (_, _, b) = bottom.color.to_rgb();
    assert!(b > 200, "Bottom border should be blue");

    cleanup_fixture(&fixture);
}

#[test]
fn test_border_on_range() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        for row in 0..3 {
            for col in 0..3 {
                let cell = format!("{}{}", (b'A' + col as u8) as char, row + 1);
                wb.set_cell_value(&cell, (row * 3 + col + 1) as f64)
                    .expect("set value");
                // xlContinuous = 1, xlMedium = -4138, black
                wb.set_border_all(&cell, 1, -4138, 0x000000)
                    .expect("set border");
            }
        }

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");

    let mut cells_with_borders = 0;
    for row in 0..3u32 {
        for col in 0..3u16 {
            if let Some(style) = sheet.cell_style_at(row, col) {
                if style.border.left.is_some()
                    || style.border.right.is_some()
                    || style.border.top.is_some()
                    || style.border.bottom.is_some()
                {
                    cells_with_borders += 1;
                }
            }
        }
    }
    assert!(
        cells_with_borders >= 9,
        "All 9 cells should have borders, got {cells_with_borders}"
    );

    cleanup_fixture(&fixture);
}
