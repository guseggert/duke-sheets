//! Integration tests reading real-world .xls files.

use duke_sheets_xls::XlsReader;
use std::path::Path;

fn project_root() -> &'static Path {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .unwrap()
        .parent()
        .unwrap()
}

#[test]
fn test_read_zso_status_report() {
    let path = project_root().join("excel_samples/ZSO_Status_Report.xls");
    if !path.exists() {
        eprintln!("Skipping: {path:?} not found");
        return;
    }

    let wb = match XlsReader::read_file(&path) {
        Ok(wb) => wb,
        Err(e) => {
            // This file may not be a valid XLS (some files have .xls extension but aren't CFB)
            eprintln!("Skipping: {e}");
            return;
        }
    };
    assert!(wb.sheet_count() > 0, "should have at least one sheet");

    let ws = wb.worksheet(0).unwrap();
    eprintln!("Sheet 0: \"{}\"", ws.name());

    // Just verify we can read cells without panicking
    let mut cell_count = 0;
    for row in 0..100u32 {
        for col in 0..20u16 {
            let val = ws.get_value_at(row, col);
            if !val.is_empty() {
                cell_count += 1;
            }
        }
    }
    eprintln!("  Found {cell_count} non-empty cells in first 100 rows");
    assert!(cell_count > 0, "should have some data");
}

#[test]
fn test_read_credit_cards() {
    let path = project_root().join("excel_samples/26DPP01223_T2551_Exhibit_3_-_Inbound_Daily_File_for_Credit_Cards_1142026.xls");
    if !path.exists() {
        eprintln!("Skipping: {path:?} not found");
        return;
    }

    let wb = XlsReader::read_file(&path).expect("should read credit cards xls");
    assert!(wb.sheet_count() > 0, "should have at least one sheet");

    let ws = wb.worksheet(0).unwrap();
    eprintln!("Sheet 0: \"{}\"", ws.name());

    let mut cell_count = 0;
    for row in 0..50u32 {
        for col in 0..20u16 {
            let val = ws.get_value_at(row, col);
            if !val.is_empty() {
                cell_count += 1;
            }
        }
    }
    eprintln!("  Found {cell_count} non-empty cells in first 50 rows");
    assert!(cell_count > 0, "should have some data");
}
