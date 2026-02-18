//! Tests for reading merged cells from XLSX files.

use crate::{cleanup_fixture, lo_bridge, runtime, skip_if_no_lo, temp_fixture_path};
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_merged_cells_horizontal() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Merged horizontal").await.unwrap();
        wb.merge_range(0, "A1:C1").await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let regions = sheet.merged_regions();
    assert_eq!(regions.len(), 1, "Should have 1 merged region, got {}", regions.len());
    let r = &regions[0];
    assert_eq!(r.start.row, 0, "Start row should be 0");
    assert_eq!(r.start.col, 0, "Start col should be 0 (A)");
    assert_eq!(r.end.row, 0, "End row should be 0");
    assert_eq!(r.end.col, 2, "End col should be 2 (C)");

    cleanup_fixture(&path);
}

#[test]
fn test_merged_cells_vertical() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Merged vertical").await.unwrap();
        wb.merge_range(0, "A1:A3").await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let regions = sheet.merged_regions();
    assert_eq!(regions.len(), 1, "Should have 1 merged region");
    let r = &regions[0];
    assert_eq!(r.start.row, 0);
    assert_eq!(r.start.col, 0);
    assert_eq!(r.end.row, 2, "End row should be 2 (row 3)");
    assert_eq!(r.end.col, 0);

    cleanup_fixture(&path);
}

#[test]
fn test_merged_cells_block() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Merged block").await.unwrap();
        wb.merge_range(0, "A1:C3").await.unwrap();
        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();
    let regions = sheet.merged_regions();
    assert_eq!(regions.len(), 1, "Should have 1 merged region");
    let r = &regions[0];
    assert_eq!(r.start.row, 0);
    assert_eq!(r.start.col, 0);
    assert_eq!(r.end.row, 2);
    assert_eq!(r.end.col, 2);

    cleanup_fixture(&path);
}
