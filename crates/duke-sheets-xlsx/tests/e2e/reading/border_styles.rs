//! Tests for reading border styles from XLSX files.

use crate::{cleanup_fixture, lo_bridge, runtime, skip_if_no_lo, temp_fixture_path};
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_thin_border_all_sides() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Thin border").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            border_style: Some("thin".to_string()),
            border_color: Some(0x000000),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.border.left.is_some(), "Should have left border");
    assert!(style.border.right.is_some(), "Should have right border");
    assert!(style.border.top.is_some(), "Should have top border");
    assert!(style.border.bottom.is_some(), "Should have bottom border");

    cleanup_fixture(&path);
}

#[test]
fn test_border_color() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Red border").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            border_style: Some("medium".to_string()),
            border_color: Some(0xFF0000),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");

    let edge = style.border.top.as_ref().or(style.border.left.as_ref()).expect("Should have a border");
    let (r, _, _) = edge.color.to_rgb();
    assert!(r > 200, "Expected red border color");

    cleanup_fixture(&path);
}

#[test]
fn test_individual_border_left_only() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Left only").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            left_border: Some(("thin".to_string(), 0x000000)),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.border.left.is_some(), "Should have left border");
    assert!(style.border.right.is_none(), "Should NOT have right border");
    assert!(style.border.top.is_none(), "Should NOT have top border");
    assert!(style.border.bottom.is_none(), "Should NOT have bottom border");

    cleanup_fixture(&path);
}

#[test]
fn test_mixed_border_sides() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Mixed").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            top_border: Some(("thin".to_string(), 0xFF0000)),
            bottom_border: Some(("thick".to_string(), 0x0000FF)),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    assert!(style.border.top.is_some(), "Should have top border");
    assert!(style.border.bottom.is_some(), "Should have bottom border");

    let top = style.border.top.as_ref().unwrap();
    let (r, _, _) = top.color.to_rgb();
    assert!(r > 200, "Top border should be red");

    let bottom = style.border.bottom.as_ref().unwrap();
    let (_, _, b) = bottom.color.to_rgb();
    assert!(b > 200, "Bottom border should be blue");

    cleanup_fixture(&path);
}

#[test]
fn test_border_on_range() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        for row in 0..3 {
            for col in 0..3 {
                let cell = format!("{}{}", (b'A' + col as u8) as char, row + 1);
                wb.set_cell_value(&cell, (row * 3 + col + 1) as f64).await.unwrap();
            }
        }
        let spec = duke_sheets_libreoffice::StyleSpec {
            border_style: Some("medium".to_string()),
            border_color: Some(0x000000),
            ..Default::default()
        };
        wb.set_range_style(0, "A1:C3", &spec).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

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
    assert!(cells_with_borders >= 9, "All 9 cells should have borders, got {cells_with_borders}");

    cleanup_fixture(&path);
}
