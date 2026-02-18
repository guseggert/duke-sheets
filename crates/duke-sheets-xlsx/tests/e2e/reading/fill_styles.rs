//! Tests for reading fill/background styles from XLSX files.

use crate::{cleanup_fixture, lo_bridge, runtime, skip_if_no_lo, temp_fixture_path};
use duke_sheets_core::FillStyle;
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_solid_fill_red() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Red fill").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            fill_color: Some(0xFF0000),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");
    match &style.fill {
        FillStyle::Solid { color } => {
            let (r, g, b) = color.to_rgb();
            assert!(r > 200 && g < 50 && b < 50, "Expected red fill, got ({r}, {g}, {b})");
        }
        other => panic!("Expected Solid fill, got {other:?}"),
    }

    cleanup_fixture(&path);
}

#[test]
fn test_multiple_fill_colors() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        for (row, (label, color)) in [("Red", 0xFF0000i32), ("Green", 0x00FF00), ("Blue", 0x0000FF)]
            .iter()
            .enumerate()
        {
            let cell = format!("A{}", row + 1);
            wb.set_cell_value(&cell, *label).await.unwrap();
            let spec = duke_sheets_libreoffice::StyleSpec {
                fill_color: Some(*color),
                ..Default::default()
            };
            wb.set_cell_style(0, &cell, &spec).await.unwrap();
        }
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

    let mut fill_count = 0;
    for row in 0..3u32 {
        if let Some(style) = sheet.cell_style_at(row, 0) {
            if !matches!(style.fill, FillStyle::None) {
                fill_count += 1;
            }
        }
    }
    assert_eq!(fill_count, 3, "Should have 3 cells with fills");

    cleanup_fixture(&path);
}

#[test]
fn test_fill_with_white_font() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "White on Blue").await.unwrap();
        let spec = duke_sheets_libreoffice::StyleSpec {
            fill_color: Some(0x0000FF),
            font_color: Some(0xFFFFFF),
            ..Default::default()
        };
        wb.set_cell_style(0, "A1", &spec).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let style = sheet.cell_style_at(0, 0).expect("A1 should have style");

    match &style.fill {
        FillStyle::Solid { color } => {
            let (_, _, b) = color.to_rgb();
            assert!(b > 200, "Expected blue fill");
        }
        other => panic!("Expected Solid fill, got {other:?}"),
    }

    let (r, g, b) = style.font.color.to_rgb();
    assert!(r > 200 && g > 200 && b > 200, "Expected white font, got ({r}, {g}, {b})");

    cleanup_fixture(&path);
}

// Note: Gradient fill E2E test is not possible because LO Calc 7.3 cells
// don't support FillStyle/FillGradient drawing properties. Gradient fill
// read/write is verified by the roundtrip test in xlsx_style_roundtrip.rs
// (test_roundtrip_gradient_fill) which uses our own writer → reader.
