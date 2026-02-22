//! Tests for reading fill/background styles from XLSX files created by Excel.

use crate::{cleanup_fixture, ensure_vm_temp_dir, excel_bridge, pull_file_from_vm, temp_fixture};
use duke_sheets_core::FillStyle;
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_solid_fill_red() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Red fill").expect("set value");
        wb.set_fill_color("A1", 0xFF0000).expect("set fill color");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    match &style.fill {
        FillStyle::Solid { color } => {
            let (r, g, b) = color.to_rgb();
            assert!(
                r > 200 && g < 50 && b < 50,
                "Expected red fill, got ({r}, {g}, {b})"
            );
        }
        other => panic!("Expected Solid fill, got {other:?}"),
    }

    cleanup_fixture(&fixture);
}

#[test]
fn test_multiple_fill_colors() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        for (row, (label, color)) in [
            ("Red", 0xFF0000i32),
            ("Green", 0x00FF00),
            ("Blue", 0x0000FF),
        ]
        .iter()
        .enumerate()
        {
            let cell = format!("A{}", row + 1);
            wb.set_cell_value(&cell, *label).expect("set value");
            wb.set_fill_color(&cell, *color as u32).expect("set fill");
        }

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");

    let mut fill_count = 0;
    for row in 0..3u32 {
        if let Some(style) = sheet.cell_style_at(row, 0) {
            if !matches!(style.fill, FillStyle::None) {
                fill_count += 1;
            }
        }
    }
    assert_eq!(fill_count, 3, "Should have 3 cells with fills");

    cleanup_fixture(&fixture);
}

#[test]
fn test_fill_with_white_font() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "White on Blue").expect("set value");
        wb.set_fill_color("A1", 0x0000FF).expect("set fill");
        wb.set_font_color("A1", 0xFFFFFF).expect("set font color");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");

    match &style.fill {
        FillStyle::Solid { color } => {
            let (_, _, b) = color.to_rgb();
            assert!(b > 200, "Expected blue fill");
        }
        other => panic!("Expected Solid fill, got {other:?}"),
    }

    let (r, g, b) = style.font.color.to_rgb();
    assert!(
        r > 200 && g > 200 && b > 200,
        "Expected white font, got ({r}, {g}, {b})"
    );

    cleanup_fixture(&fixture);
}
