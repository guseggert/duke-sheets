#![allow(clippy::approx_constant)]
//! Integration tests for the LibreOffice URP bridge.
//!
//! These tests need a LibreOffice URP listener on localhost:2002.
//! `duke_sheets_test_harness::lo::ensure_lo()` is called at the top of
//! every test and auto-starts the docker container if it isn't already
//! running. There is no silent-skip path: if the container can't be
//! brought up the test panics.

use duke_sheets::prelude::*;
use duke_sheets_test_harness::lo::ensure_lo;

#[tokio::test]
async fn test_connect_and_bootstrap() {
    ensure_lo();

    let bridge = duke_sheets_libreoffice::LibreOfficeBridge::connect("localhost", 2002).await;
    match bridge {
        Ok(bridge) => {
            eprintln!("OK: Connected and bootstrapped successfully");
            let _ = bridge.shutdown().await;
        }
        Err(e) => {
            panic!("Failed to connect and bootstrap: {e}");
        }
    }
}

#[tokio::test]
async fn test_create_workbook_and_set_cells() {
    ensure_lo();

    let mut bridge = duke_sheets_libreoffice::LibreOfficeBridge::connect("localhost", 2002)
        .await
        .expect("connect");

    let mut wb = bridge.create_workbook().await.expect("create_workbook");

    // Set a numeric value
    wb.set_cell_value("A1", 42.0).await.expect("set A1");

    // Set a string value
    wb.set_cell_value("B1", "Hello").await.expect("set B1");

    // Set a formula
    wb.set_cell_formula("C1", "=A1*2")
        .await
        .expect("set C1 formula");

    // Read back the numeric value
    let val = wb.get_cell_value("A1").await.expect("get A1");
    assert!((val - 42.0).abs() < 1e-10, "A1 should be 42.0, got {val}");

    // Read back the formula
    let formula = wb.get_cell_formula("C1").await.expect("get C1 formula");
    assert!(
        formula.contains("A1") && formula.contains("2"),
        "C1 formula should reference A1*2, got: {formula}"
    );

    // Read back the string
    let s = wb.get_cell_string("B1").await.expect("get B1 string");
    assert_eq!(s, "Hello", "B1 should be 'Hello', got '{s}'");

    wb.close().await.expect("close");
    let _ = bridge.shutdown().await;

    eprintln!("OK: Created workbook, set cells, read back values");
}

#[tokio::test]
async fn test_save_and_read_back_with_duke_sheets() {
    ensure_lo();

    let mut bridge = duke_sheets_libreoffice::LibreOfficeBridge::connect("localhost", 2002)
        .await
        .expect("connect");

    let mut wb = bridge.create_workbook().await.expect("create_workbook");

    // Set various cell types
    wb.set_cell_value("A1", 3.14).await.expect("set A1");
    wb.set_cell_value("A2", "Hello World")
        .await
        .expect("set A2");
    wb.set_cell_formula("A3", "=A1*10").await.expect("set A3");
    wb.set_cell_value("B1", 100.0).await.expect("set B1");
    wb.set_cell_value("B2", 200.0).await.expect("set B2");
    wb.set_cell_formula("B3", "=SUM(B1:B2)")
        .await
        .expect("set B3");

    // Save to a temp file in the shared volume directory (accessible by both host and
    // the LibreOffice Docker container via -v /tmp/duke-sheets-urp:/tmp/duke-sheets-urp)
    let shared_dir = std::path::PathBuf::from("/tmp/duke-sheets-urp");
    std::fs::create_dir_all(&shared_dir).ok();
    let path = shared_dir.join(format!("urp-test-{}.xlsx", std::process::id()));
    let path_str = path.to_str().unwrap();

    wb.save(path_str).await.expect("save");
    wb.close().await.expect("close");
    let _ = bridge.shutdown().await;

    // Now read back with duke-sheets (our XLSX reader)
    assert!(path.exists(), "Saved file should exist at {path_str}");

    let file_size = std::fs::metadata(&path).unwrap().len();
    assert!(file_size > 0, "Saved file should not be empty");
    eprintln!("Saved XLSX file: {path_str} ({file_size} bytes)");

    let workbook = Workbook::open(path_str).expect("open with duke-sheets");
    let sheet = workbook.worksheet(0).expect("first sheet");

    // Check numeric value - A1 is row=0, col=0
    let a1 = sheet.get_value_at(0, 0);
    match a1 {
        CellValue::Number(n) => {
            assert!((n - 3.14).abs() < 1e-10, "A1 should be 3.14, got {n}");
        }
        other => panic!("A1 should be Number(3.14), got {other:?}"),
    }

    // Check string value - A2 is row=1, col=0
    let a2 = sheet.get_value_at(1, 0);
    match &a2 {
        CellValue::String(s) => {
            assert_eq!(
                s.as_str(),
                "Hello World",
                "A2 should be 'Hello World', got '{s}'"
            );
        }
        other => panic!("A2 should be String('Hello World'), got {other:?}"),
    }

    // Check that the formulas exist (they may be stored as computed values or formulas)
    // A3 = A1*10 = 31.4, row=2, col=0
    let a3 = sheet.get_value_at(2, 0);
    assert!(
        sheet.get_formula_at(2, 0).is_some(),
        "A3 should retain its formula"
    );
    match &a3 {
        CellValue::Number(n) => {
            assert!(
                (n - 31.4).abs() < 1e-10,
                "A3 should be 31.4 (=A1*10), got {n}"
            );
        }
        other => {
            eprintln!("A3 value: {other:?} (may vary depending on LO behavior)");
        }
    }

    // B3 = SUM(B1:B2) = 300, row=2, col=1
    let b3 = sheet.get_value_at(2, 1);
    assert!(
        sheet.get_formula_at(2, 1).is_some(),
        "B3 should retain its formula"
    );
    match &b3 {
        CellValue::Number(n) => {
            assert!(
                (n - 300.0).abs() < 1e-10,
                "B3 should be 300.0 (=SUM(B1:B2)), got {n}"
            );
        }
        other => {
            eprintln!("B3 value: {other:?}");
        }
    }

    // Clean up
    let _ = std::fs::remove_file(&path);

    eprintln!("OK: Saved XLSX via URP, read back with duke-sheets, values match");
}

#[tokio::test]
async fn test_set_array_formula() {
    ensure_lo();

    let mut bridge = duke_sheets_libreoffice::LibreOfficeBridge::connect("localhost", 2002)
        .await
        .expect("connect");
    let mut wb = bridge.create_workbook().await.expect("create_workbook");

    wb.set_cell_value("A1", 1.0).await.unwrap();
    wb.set_cell_value("A2", 2.0).await.unwrap();
    wb.set_cell_value("B1", 10.0).await.unwrap();
    wb.set_cell_value("B2", 20.0).await.unwrap();

    // Enter a CSE array formula via set_array_formula
    wb.set_array_formula(0, "C1", "=SUM(A1:A2*B1:B2)")
        .await
        .expect("set_array_formula");

    // Read back - LO wraps array formulas in braces
    let formula = wb.get_cell_formula("C1").await.expect("get_cell_formula");
    assert_eq!(formula, "{=SUM(A1:A2*B1:B2)}");

    wb.close().await.unwrap();
}
