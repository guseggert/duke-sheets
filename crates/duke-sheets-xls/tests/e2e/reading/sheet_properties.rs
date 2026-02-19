//! Tests for reading sheet-level properties from XLS files.

use crate::{cleanup_fixture, lo_bridge, runtime, skip_if_no_lo, temp_fixture_path};
use duke_sheets_xls::XlsReader;

#[test]
fn test_xls_hidden_sheet() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();

        // Sheet 0 (default) — visible
        wb.set_cell_value("A1", "Visible sheet").await.unwrap();

        // Add sheet 1 — will be hidden
        wb.add_sheet("Hidden").await.unwrap();
        let cell = wb.get_cell_on_sheet(1, 0, 0).await.unwrap();
        wb.set_cell_value_on_proxy(&cell, "Hidden sheet data".into())
            .await
            .unwrap();
        wb.set_sheet_hidden(1, true).await.unwrap();

        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();

    let sheet0 = workbook.worksheet(0).unwrap();
    assert!(sheet0.is_visible(), "Sheet 0 should be visible");

    let sheet1 = workbook.worksheet(1).unwrap();
    assert_eq!(sheet1.name(), "Hidden");
    assert!(!sheet1.is_visible(), "Sheet 1 should be hidden");

    cleanup_fixture(&path);
}

#[test]
fn test_xls_active_sheet() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();

        wb.set_cell_value("A1", "Sheet 1 data").await.unwrap();

        wb.add_sheet("Second").await.unwrap();
        let cell = wb.get_cell_on_sheet(1, 0, 0).await.unwrap();
        wb.set_cell_value_on_proxy(&cell, "Sheet 2 data".into())
            .await
            .unwrap();

        wb.add_sheet("Third").await.unwrap();
        let cell = wb.get_cell_on_sheet(2, 0, 0).await.unwrap();
        wb.set_cell_value_on_proxy(&cell, "Sheet 3 data".into())
            .await
            .unwrap();

        // Make sheet 1 (zero-indexed) the active sheet
        wb.set_active_sheet(1).await.unwrap();

        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    assert_eq!(workbook.active_sheet(), 1, "Active sheet should be 1");
    assert_eq!(workbook.sheet_count(), 3);

    cleanup_fixture(&path);
}

#[test]
fn test_xls_sheet_protection() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();

        wb.set_cell_value("A1", "Protected sheet").await.unwrap();

        // Add a second sheet (unprotected)
        wb.add_sheet("Unprotected").await.unwrap();
        let cell = wb.get_cell_on_sheet(1, 0, 0).await.unwrap();
        wb.set_cell_value_on_proxy(&cell, "Open sheet".into())
            .await
            .unwrap();

        // Protect sheet 0 with an empty password
        wb.protect_sheet(0, "").await.unwrap();

        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();

    let sheet0 = workbook.worksheet(0).unwrap();
    let prot = sheet0.protection();
    assert!(
        prot.is_some(),
        "Sheet 0 should have protection settings"
    );
    assert!(
        prot.unwrap().protected,
        "Sheet 0 should be protected"
    );

    let sheet1 = workbook.worksheet(1).unwrap();
    assert!(
        sheet1.protection().is_none(),
        "Sheet 1 should not be protected"
    );

    cleanup_fixture(&path);
}
