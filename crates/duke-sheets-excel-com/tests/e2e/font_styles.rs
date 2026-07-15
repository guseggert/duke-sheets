//! Tests for reading font styles from XLSX files created by Excel.

use crate::{cleanup_fixture, ensure_vm_temp_dir, excel_bridge, pull_file_from_vm, temp_fixture};
use duke_sheets_core::style::{FontVerticalAlign, Underline};
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_bold() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Bold").expect("set value");
        wb.set_font_bold("A1", true).expect("set bold");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.font.bold, "Font should be bold");

    cleanup_fixture(&fixture);
}

#[test]
fn test_italic() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Italic").expect("set value");
        wb.set_font_italic("A1", true).expect("set italic");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.font.italic, "Font should be italic");

    cleanup_fixture(&fixture);
}

#[test]
fn test_superscript() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Super").expect("set value");
        wb.set_font_superscript("A1", true)
            .expect("set superscript");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(
        style.font.vertical_align,
        FontVerticalAlign::Superscript,
        "Font should be superscript"
    );

    cleanup_fixture(&fixture);
}

#[test]
fn test_subscript() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Sub").expect("set value");
        wb.set_font_subscript("A1", true).expect("set subscript");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(
        style.font.vertical_align,
        FontVerticalAlign::Subscript,
        "Font should be subscript"
    );

    cleanup_fixture(&fixture);
}

#[test]
fn test_underline_single() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Underline").expect("set value");
        // xlUnderlineStyleSingle = 2
        wb.set_font_underline("A1", 2).expect("set underline");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(style.font.underline, Underline::Single);

    cleanup_fixture(&fixture);
}

#[test]
fn test_strikethrough() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Strike").expect("set value");
        wb.set_font_strikethrough("A1", true)
            .expect("set strikethrough");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.font.strikethrough, "Font should be strikethrough");

    cleanup_fixture(&fixture);
}

#[test]
fn test_font_color() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Red text").expect("set value");
        wb.set_font_color("A1", 0xFF0000).expect("set color");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    let (r, g, b) = style.font.color.to_rgb().unwrap();
    assert!(
        r > 200 && g < 50 && b < 50,
        "Expected red font, got ({r}, {g}, {b})"
    );

    cleanup_fixture(&fixture);
}

#[test]
fn test_font_size() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Big").expect("set value");
        wb.set_font_size("A1", 20.0).expect("set size");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(
        (style.font.size - 20.0).abs() < 0.5,
        "Expected font size ~20, got {}",
        style.font.size
    );

    cleanup_fixture(&fixture);
}

#[test]
fn test_font_name() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Courier").expect("set value");
        wb.set_font_name("A1", "Courier New").expect("set name");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(style.font.name, "Courier New");

    cleanup_fixture(&fixture);
}

#[test]
fn test_font_style_combination() {
    let bridge = excel_bridge();
    let fixture = temp_fixture();

    {
        let excel = bridge.lock().unwrap();
        ensure_vm_temp_dir();
        let wb = excel.create_workbook().expect("create workbook");

        wb.set_cell_value("A1", "Combo").expect("set value");
        wb.set_font_bold("A1", true).expect("set bold");
        wb.set_font_italic("A1", true).expect("set italic");
        wb.set_font_underline("A1", 2).expect("set underline"); // xlUnderlineStyleSingle
        wb.set_font_color("A1", 0x0000FF).expect("set color"); // blue
        wb.set_font_size("A1", 14.0).expect("set size");

        wb.save(&fixture.vm_path).expect("save");
        wb.close().expect("close");
    }

    pull_file_from_vm(&fixture);
    let workbook = XlsxReader::read_file(&fixture.host_path).expect("read");
    let sheet = workbook.worksheet(0).expect("worksheet");
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");

    assert!(style.font.bold, "Should be bold");
    assert!(style.font.italic, "Should be italic");
    assert_eq!(style.font.underline, Underline::Single);
    let (r, g, b) = style.font.color.to_rgb().unwrap();
    assert!(b > 200 && r < 50, "Should be blue, got ({r}, {g}, {b})");
    assert!(
        (style.font.size - 14.0).abs() < 0.5,
        "Expected size ~14, got {}",
        style.font.size
    );

    cleanup_fixture(&fixture);
}
