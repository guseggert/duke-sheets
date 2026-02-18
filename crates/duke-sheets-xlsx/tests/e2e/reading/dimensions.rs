//! Tests for reading row heights and column widths from XLSX files.

use crate::{cleanup_fixture, lo_bridge, runtime, skip_if_no_lo, temp_fixture_path};
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_custom_row_height() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Tall row").await.unwrap();
        wb.set_row_height(0, 0, 30.0).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let height = sheet.row_height(0);
    assert!(
        (height - 30.0).abs() < 1.5,
        "Row height should be ~30, got {}",
        height
    );

    cleanup_fixture(&path);
}

#[test]
fn test_custom_column_width() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Wide column").await.unwrap();
        wb.set_column_width(0, 0, 20.0).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let width = sheet.column_width(0);
    // LO's column width conversion is approximate — allow some tolerance
    assert!(
        width > 15.0,
        "Column width should be significantly wider than default (8.43), got {}",
        width
    );

    cleanup_fixture(&path);
}

#[test]
fn test_multiple_row_heights() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Row 1").await.unwrap();
        wb.set_cell_value("A2", "Row 2").await.unwrap();
        wb.set_cell_value("A3", "Row 3").await.unwrap();
        wb.set_row_height(0, 0, 25.0).await.unwrap();
        wb.set_row_height(0, 2, 40.0).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let custom = sheet.custom_row_heights();
    assert!(
        custom.len() >= 2,
        "Should have at least 2 custom row heights, got {}",
        custom.len()
    );

    cleanup_fixture(&path);
}

#[test]
fn test_hidden_row() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Visible").await.unwrap();
        wb.set_cell_value("A2", "Hidden").await.unwrap();
        wb.set_cell_value("A3", "Visible").await.unwrap();
        wb.set_row_hidden(0, 1, true).await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    assert!(!sheet.is_row_hidden(0), "Row 0 should not be hidden");
    assert!(sheet.is_row_hidden(1), "Row 1 should be hidden");
    assert!(!sheet.is_row_hidden(2), "Row 2 should not be hidden");

    cleanup_fixture(&path);
}
