//! Tests for reading font styles from XLSX files.

use crate::{cleanup_fixture, lo_bridge, runtime, skip_if_no_lo, temp_fixture_path};
use duke_sheets_core::style::{FontVerticalAlign, Underline};
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_bold() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Bold").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            bold: true,
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.font.bold, "Font should be bold");

    cleanup_fixture(&path);
}

#[test]
fn test_superscript() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Super").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            font_vertical_align: Some("superscript".to_string()),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(
        style.font.vertical_align,
        FontVerticalAlign::Superscript,
        "Font should be superscript"
    );

    cleanup_fixture(&path);
}

#[test]
fn test_subscript() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Sub").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            font_vertical_align: Some("subscript".to_string()),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(
        style.font.vertical_align,
        FontVerticalAlign::Subscript,
        "Font should be subscript"
    );

    cleanup_fixture(&path);
}


#[test]
fn test_underline_single() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Underline").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            underline: Some("single".to_string()),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(style.font.underline, Underline::Single);

    cleanup_fixture(&path);
}

#[test]
fn test_strikethrough() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Strike").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            strikethrough: true,
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.font.strikethrough, "Font should be strikethrough");

    cleanup_fixture(&path);
}

#[test]
fn test_font_color() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Red text").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            font_color: Some(0xFF0000i32),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    let (r, g, b) = style.font.color.to_rgb();
    assert!(r > 200 && g < 50 && b < 50, "Expected red font, got ({r}, {g}, {b})");

    cleanup_fixture(&path);
}

#[test]
fn test_font_size() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Big").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            font_size: Some(20.0),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(
        (style.font.size - 20.0).abs() < 0.5,
        "Expected font size ~20, got {}",
        style.font.size
    );

    cleanup_fixture(&path);
}

#[test]
fn test_font_name() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Courier").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            font_name: Some("Courier New".to_string()),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert_eq!(style.font.name, "Courier New");

    cleanup_fixture(&path);
}

#[test]
fn test_font_style_combination() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Combo").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            bold: true,
            italic: true,
            underline: Some("single".to_string()),
            font_color: Some(0x0000FF),
            font_size: Some(14.0),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.font.bold, "Should be bold");
    assert!(style.font.italic, "Should be italic");
    assert_eq!(style.font.underline, Underline::Single);
    let (r, g, b) = style.font.color.to_rgb();
    assert!(b > 200 && r < 50, "Should be blue, got ({r}, {g}, {b})");
    assert!(
        (style.font.size - 14.0).abs() < 0.5,
        "Expected size ~14, got {}",
        style.font.size
    );

    cleanup_fixture(&path);
}
