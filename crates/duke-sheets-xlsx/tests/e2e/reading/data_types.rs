//! Tests for reading data types from XLSX files.
//!
//! Each test creates its fixture on-demand via LibreOffice, saves to a temp
//! file, reads it back with `XlsxReader`, and asserts.

use crate::{cleanup_fixture, lo_bridge, runtime, skip_if_no_lo, temp_fixture_path};
use duke_sheets_xlsx::XlsxReader;

#[test]
fn test_number_values() {
    skip_if_no_lo!();

    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.expect("LO should be available");
        let mut bridge = lo.lock().await;
        let mut wb = bridge.create_workbook().await.expect("create workbook");

        wb.set_cell_value("A1", 42.0).await.unwrap();
        wb.set_cell_value("A2", 3.14159).await.unwrap();
        wb.set_cell_value("A3", -100.0).await.unwrap();
        wb.set_cell_value("A4", 0.0).await.unwrap();

        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).expect("Failed to read workbook");
    let sheet = workbook.worksheet(0).expect("No worksheet");

    let cell = sheet.cell_at(0, 0).expect("A1 should exist");
    match &cell.value {
        duke_sheets_core::CellValue::Number(n) => {
            assert!((*n - 42.0).abs() < 0.001, "Expected 42, got {n}");
        }
        other => panic!("Expected Number, got {other:?}"),
    }

    let cell = sheet.cell_at(1, 0).expect("A2 should exist");
    match &cell.value {
        duke_sheets_core::CellValue::Number(n) => {
            assert!((*n - 3.14159).abs() < 0.001, "Expected 3.14159, got {n}");
        }
        other => panic!("Expected Number, got {other:?}"),
    }

    let cell = sheet.cell_at(2, 0).expect("A3 should exist");
    match &cell.value {
        duke_sheets_core::CellValue::Number(n) => {
            assert!((*n + 100.0).abs() < 0.001, "Expected -100, got {n}");
        }
        other => panic!("Expected Number, got {other:?}"),
    }

    let cell = sheet.cell_at(3, 0).expect("A4 should exist");
    match &cell.value {
        duke_sheets_core::CellValue::Number(n) => {
            assert!(n.abs() < 0.001, "Expected 0, got {n}");
        }
        other => panic!("Expected Number, got {other:?}"),
    }

    cleanup_fixture(&path);
}

#[test]
fn test_string_values() {
    skip_if_no_lo!();

    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.expect("LO should be available");
        let mut bridge = lo.lock().await;
        let mut wb = bridge.create_workbook().await.expect("create workbook");

        wb.set_cell_value("A1", "Hello").await.unwrap();
        wb.set_cell_value("A2", "World with spaces").await.unwrap();
        wb.set_cell_value("A3", "Unicode: \u{65e5}\u{672c}\u{8a9e}").await.unwrap();

        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).expect("Failed to read workbook");
    let sheet = workbook.worksheet(0).expect("No worksheet");

    let cell = sheet.cell_at(0, 0).expect("A1 should exist");
    match &cell.value {
        duke_sheets_core::CellValue::String(s) => {
            assert_eq!(s.as_ref(), "Hello");
        }
        other => panic!("Expected String, got {other:?}"),
    }

    let cell = sheet.cell_at(1, 0).expect("A2 should exist");
    match &cell.value {
        duke_sheets_core::CellValue::String(s) => {
            assert_eq!(s.as_ref(), "World with spaces");
        }
        other => panic!("Expected String, got {other:?}"),
    }

    let cell = sheet.cell_at(2, 0).expect("A3 should exist");
    match &cell.value {
        duke_sheets_core::CellValue::String(s) => {
            assert!(
                s.as_ref().contains("\u{65e5}\u{672c}\u{8a9e}"),
                "Expected Japanese text, got {s}"
            );
        }
        other => panic!("Expected String, got {other:?}"),
    }

    cleanup_fixture(&path);
}

#[test]
fn test_boolean_values() {
    skip_if_no_lo!();

    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.expect("LO should be available");
        let mut bridge = lo.lock().await;
        let mut wb = bridge.create_workbook().await.expect("create workbook");

        // Use =TRUE() and =FALSE() as formulas so LO preserves them
        // (plain "TRUE" gets optimized to a numeric constant on save)
        wb.set_cell_formula("A1", "=TRUE()").await.unwrap();
        wb.set_cell_formula("A2", "=FALSE()").await.unwrap();

        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).expect("Failed to read workbook");
    let sheet = workbook.worksheet(0).expect("No worksheet");

    // Use effective_value() to look through formulas to their cached values
    let mut found_true = false;
    let mut found_false = false;
    for row in 0..5 {
        for col in 0..5 {
            if let Some(cell) = sheet.cell_at(row, col) {
                match cell.value.effective_value() {
                    duke_sheets_core::CellValue::Boolean(b) => {
                        if *b {
                            found_true = true;
                        } else {
                            found_false = true;
                        }
                    }
                    // LO may also cache boolean results as numbers (1/0)
                    duke_sheets_core::CellValue::Number(n) => {
                        if (*n - 1.0).abs() < 0.001 {
                            found_true = true;
                        } else if n.abs() < 0.001 {
                            found_false = true;
                        }
                    }
                    _ => {}
                }
            }
        }
    }

    assert!(found_true, "Should find TRUE value");
    assert!(found_false, "Should find FALSE value");

    cleanup_fixture(&path);
}

#[test]
fn test_formula_values() {
    skip_if_no_lo!();

    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.expect("LO should be available");
        let mut bridge = lo.lock().await;
        let mut wb = bridge.create_workbook().await.expect("create workbook");

        wb.set_cell_value("A1", 10.0).await.unwrap();
        wb.set_cell_value("A2", 20.0).await.unwrap();
        wb.set_cell_formula("A3", "=A1+A2").await.unwrap();
        wb.set_cell_formula("A4", "=SUM(A1:A2)").await.unwrap();

        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).expect("Failed to read workbook");
    let sheet = workbook.worksheet(0).expect("No worksheet");

    let formula = sheet.get_formula_at(2, 0);
    assert!(formula.is_some(), "A3 should have a formula");

    let formula = sheet.get_formula_at(3, 0);
    assert!(formula.is_some(), "A4 should have a formula");

    cleanup_fixture(&path);
}

#[test]
fn test_error_values() {
    skip_if_no_lo!();

    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.expect("LO should be available");
        let mut bridge = lo.lock().await;
        let mut wb = bridge.create_workbook().await.expect("create workbook");

        wb.set_cell_formula("A1", "=1/0").await.unwrap(); // #DIV/0!
        wb.set_cell_formula("A2", "=VALUE(\"x\")").await.unwrap(); // #VALUE!

        wb.save(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsxReader::read_file(&path).expect("Failed to read workbook");
    let sheet = workbook.worksheet(0).expect("No worksheet");

    let mut found_error = false;
    for row in 0..5 {
        for col in 0..5 {
            if let Some(cell) = sheet.cell_at(row, col) {
                if let duke_sheets_core::CellValue::Error(_) = cell.value.effective_value() {
                    found_error = true;
                    break;
                }
            }
        }
        if found_error {
            break;
        }
    }

    assert!(found_error, "Should find at least one error value");

    cleanup_fixture(&path);
}
