//! Tests for reading various cell data types from XLS files.

use crate::{cleanup_fixture, lo_bridge, runtime, skip_if_no_lo, temp_fixture_path};
use duke_sheets_core::CellValue;
use duke_sheets_xls::XlsReader;

#[test]
fn test_xls_number_values() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", 42.0).await.unwrap();
        wb.set_cell_value("B1", 3.14).await.unwrap();
        wb.set_cell_value("C1", -100.0).await.unwrap();
        wb.set_cell_value("D1", 0.0).await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

    let a1 = sheet.get_value_at(0, 0);
    assert!(
        matches!(a1, CellValue::Number(n) if (n - 42.0).abs() < f64::EPSILON),
        "A1 should be 42.0, got {a1:?}"
    );

    let b1 = sheet.get_value_at(0, 1);
    assert!(
        matches!(b1, CellValue::Number(n) if (n - 3.14).abs() < 0.001),
        "B1 should be ~3.14, got {b1:?}"
    );

    let c1 = sheet.get_value_at(0, 2);
    assert!(
        matches!(c1, CellValue::Number(n) if (n - (-100.0)).abs() < f64::EPSILON),
        "C1 should be -100, got {c1:?}"
    );

    let d1 = sheet.get_value_at(0, 3);
    assert!(
        matches!(d1, CellValue::Number(n) if n.abs() < f64::EPSILON),
        "D1 should be 0, got {d1:?}"
    );

    cleanup_fixture(&path);
}

#[test]
fn test_xls_string_values() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "Hello").await.unwrap();
        wb.set_cell_value("B1", "World").await.unwrap();
        wb.set_cell_value("C1", "").await.unwrap();
        wb.set_cell_value("D1", "Special chars: <>&\"'")
            .await
            .unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

    let a1 = sheet.get_value("A1").unwrap();
    assert_eq!(a1.as_string(), Some("Hello"), "A1 should be 'Hello'");

    let b1 = sheet.get_value("B1").unwrap();
    assert_eq!(b1.as_string(), Some("World"), "B1 should be 'World'");

    let d1 = sheet.get_value("D1").unwrap();
    assert_eq!(
        d1.as_string(),
        Some("Special chars: <>&\"'"),
        "D1 should preserve special chars"
    );

    cleanup_fixture(&path);
}

#[test]
fn test_xls_boolean_values() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        // LO doesn't have a direct "set boolean" — use formulas
        wb.set_cell_formula("A1", "=TRUE()").await.unwrap();
        wb.set_cell_formula("B1", "=FALSE()").await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

    // TRUE()/FALSE() formulas have boolean cached results
    let a1 = sheet.get_value_at(0, 0);
    match &a1 {
        CellValue::Boolean(true) => {}
        CellValue::Formula { cached_value: Some(cv), .. } => {
            assert!(
                matches!(cv.as_ref(), CellValue::Boolean(true)),
                "A1 formula should cache TRUE, got {cv:?}"
            );
        }
        other => panic!("A1 should be TRUE or formula with TRUE cache, got {other:?}"),
    }

    let b1 = sheet.get_value_at(0, 1);
    match &b1 {
        CellValue::Boolean(false) => {}
        CellValue::Formula { cached_value: Some(cv), .. } => {
            assert!(
                matches!(cv.as_ref(), CellValue::Boolean(false)),
                "B1 formula should cache FALSE, got {cv:?}"
            );
        }
        other => panic!("B1 should be FALSE or formula with FALSE cache, got {other:?}"),
    }

    cleanup_fixture(&path);
}

#[test]
fn test_xls_formula_values() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", 10.0).await.unwrap();
        wb.set_cell_value("B1", 20.0).await.unwrap();
        wb.set_cell_formula("C1", "=A1+B1").await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

    let a1 = sheet.get_value_at(0, 0);
    assert!(
        matches!(a1, CellValue::Number(n) if (n - 10.0).abs() < f64::EPSILON),
        "A1 should be 10, got {a1:?}"
    );

    let c1 = sheet.get_value_at(0, 2);
    match &c1 {
        CellValue::Formula {
            cached_value: Some(cv),
            ..
        } => {
            assert!(
                matches!(cv.as_ref(), CellValue::Number(n) if (*n - 30.0).abs() < f64::EPSILON),
                "C1 formula cached value should be 30, got {cv:?}"
            );
        }
        other => panic!("C1 should be a formula with cached value 30, got {other:?}"),
    }

    cleanup_fixture(&path);
}

#[test]
fn test_xls_error_values() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_formula("A1", "=1/0").await.unwrap(); // #DIV/0!
        wb.set_cell_formula("B1", "=NA()").await.unwrap(); // #N/A
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

    let a1 = sheet.get_value_at(0, 0);
    match &a1 {
        CellValue::Error(e) => {
            assert!(
                matches!(e, duke_sheets_core::CellError::Div0),
                "A1 should be #DIV/0!, got {e:?}"
            );
        }
        CellValue::Formula { cached_value: Some(cv), .. } => {
            assert!(
                matches!(cv.as_ref(), CellValue::Error(duke_sheets_core::CellError::Div0)),
                "A1 formula cached value should be #DIV/0!, got {cv:?}"
            );
        }
        other => panic!("A1 should be #DIV/0!, got {other:?}"),
    }

    let b1 = sheet.get_value_at(0, 1);
    match &b1 {
        CellValue::Error(e) => {
            assert!(
                matches!(e, duke_sheets_core::CellError::Na),
                "B1 should be #N/A, got {e:?}"
            );
        }
        CellValue::Formula { cached_value: Some(cv), .. } => {
            assert!(
                matches!(cv.as_ref(), CellValue::Error(duke_sheets_core::CellError::Na)),
                "B1 formula cached value should be #N/A, got {cv:?}"
            );
        }
        other => panic!("B1 should be #N/A, got {other:?}"),
    }

    cleanup_fixture(&path);
}

#[test]
fn test_xls_multiple_sheets() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_sheet_name(0, "First").await.unwrap();
        wb.set_cell_value("A1", "Sheet1 Data").await.unwrap();

        wb.add_sheet("Second").await.unwrap();
        let cell = wb.get_cell_on_sheet(1, 0, 0).await.unwrap();
        wb.set_cell_value_on_proxy(&cell, duke_sheets_libreoffice::CellValue::String("Sheet2 Data".into()))
            .await
            .unwrap();

        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    assert!(workbook.sheet_count() >= 2, "Should have at least 2 sheets");

    let sheet1 = workbook.worksheet(0).unwrap();
    assert_eq!(sheet1.name(), "First");
    let val1 = sheet1.get_value("A1").unwrap();
    assert_eq!(val1.as_string(), Some("Sheet1 Data"));

    let sheet2 = workbook.worksheet(1).unwrap();
    assert_eq!(sheet2.name(), "Second");
    let val2 = sheet2.get_value("A1").unwrap();
    assert_eq!(val2.as_string(), Some("Sheet2 Data"));

    cleanup_fixture(&path);
}
