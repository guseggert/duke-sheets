//! E2E tests for BIFF8 formula token decompilation.
//!
//! Each test creates a workbook in LibreOffice with specific formulas,
//! saves as .xls, then reads back and verifies the decompiled formula text.

use crate::{cleanup_fixture, lo_bridge, runtime, skip_if_no_lo, temp_fixture_path};
use duke_sheets_core::CellValue;
use duke_sheets_xls::XlsReader;

/// Helper to extract formula text from a cell value.
fn formula_text(val: &CellValue) -> &str {
    match val {
        CellValue::Formula { text, .. } => text.as_str(),
        other => panic!("Expected Formula, got {:?}", other),
    }
}

#[test]
fn test_xls_formula_simple_arithmetic() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", 10.0).await.unwrap();
        wb.set_cell_value("B1", 20.0).await.unwrap();
        // C1 = A1+B1
        wb.set_cell_formula("C1", "=A1+B1").await.unwrap();
        // D1 = A1*B1
        wb.set_cell_formula("D1", "=A1*B1").await.unwrap();
        // E1 = A1-B1
        wb.set_cell_formula("E1", "=A1-B1").await.unwrap();
        // F1 = A1/B1
        wb.set_cell_formula("F1", "=A1/B1").await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

    let c1 = sheet.get_value_at(0, 2);
    assert_eq!(formula_text(&c1), "=A1+B1", "C1 formula");

    let d1 = sheet.get_value_at(0, 3);
    assert_eq!(formula_text(&d1), "=A1*B1", "D1 formula");

    let e1 = sheet.get_value_at(0, 4);
    assert_eq!(formula_text(&e1), "=A1-B1", "E1 formula");

    let f1 = sheet.get_value_at(0, 5);
    assert_eq!(formula_text(&f1), "=A1/B1", "F1 formula");

    cleanup_fixture(&path);
}

#[test]
fn test_xls_formula_functions() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        // Set up some data
        for i in 1..=5 {
            wb.set_cell_value(&format!("A{i}"), i as f64)
                .await
                .unwrap();
        }
        // B1 = SUM(A1:A5)
        wb.set_cell_formula("B1", "=SUM(A1:A5)").await.unwrap();
        // B2 = AVERAGE(A1:A5)
        wb.set_cell_formula("B2", "=AVERAGE(A1:A5)").await.unwrap();
        // B3 = MAX(A1:A5)
        wb.set_cell_formula("B3", "=MAX(A1:A5)").await.unwrap();
        // B4 = MIN(A1:A5)
        wb.set_cell_formula("B4", "=MIN(A1:A5)").await.unwrap();
        // B5 = COUNT(A1:A5)
        wb.set_cell_formula("B5", "=COUNT(A1:A5)").await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

    assert_eq!(formula_text(&sheet.get_value_at(0, 1)), "=SUM(A1:A5)");
    assert_eq!(formula_text(&sheet.get_value_at(1, 1)), "=AVERAGE(A1:A5)");
    assert_eq!(formula_text(&sheet.get_value_at(2, 1)), "=MAX(A1:A5)");
    assert_eq!(formula_text(&sheet.get_value_at(3, 1)), "=MIN(A1:A5)");
    assert_eq!(formula_text(&sheet.get_value_at(4, 1)), "=COUNT(A1:A5)");

    cleanup_fixture(&path);
}

#[test]
fn test_xls_formula_if_and_comparison() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    // LO uses semicolons as argument separators (locale-dependent), but
    // BIFF8 token streams are locale-independent — the decompiler always
    // produces commas.
    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", 10.0).await.unwrap();
        // B1 = IF(A1>5;1;0)
        wb.set_cell_formula("B1", "=IF(A1>5;1;0)").await.unwrap();
        // C1 = IF(A1<>0;"yes";"no")
        wb.set_cell_formula("C1", "=IF(A1<>0;\"yes\";\"no\")")
            .await
            .unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

    assert_eq!(formula_text(&sheet.get_value_at(0, 1)), "=IF(A1>5,1,0)");
    assert_eq!(
        formula_text(&sheet.get_value_at(0, 2)),
        "=IF(A1<>0,\"yes\",\"no\")"
    );

    cleanup_fixture(&path);
}

#[test]
fn test_xls_formula_string_concat() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "hello").await.unwrap();
        // B1 = CONCATENATE(A1;" world") — use function form for locale safety
        wb.set_cell_formula("B1", "=CONCATENATE(A1;\" world\")")
            .await
            .unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

    assert_eq!(
        formula_text(&sheet.get_value_at(0, 1)),
        "=CONCATENATE(A1,\" world\")"
    );

    cleanup_fixture(&path);
}

#[test]
fn test_xls_formula_nested_functions() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", "hello").await.unwrap();
        // B1 = LEN(A1)
        wb.set_cell_formula("B1", "=LEN(A1)").await.unwrap();
        // C1 = LEFT(A1;3) — semicolon for LO locale
        wb.set_cell_formula("C1", "=LEFT(A1;3)").await.unwrap();
        // D1 = MID(A1;2;3) — semicolons for LO locale
        wb.set_cell_formula("D1", "=MID(A1;2;3)").await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

    assert_eq!(formula_text(&sheet.get_value_at(0, 1)), "=LEN(A1)");
    assert_eq!(formula_text(&sheet.get_value_at(0, 2)), "=LEFT(A1,3)");
    assert_eq!(formula_text(&sheet.get_value_at(0, 3)), "=MID(A1,2,3)");

    cleanup_fixture(&path);
}

#[test]
fn test_xls_formula_constants_and_unary() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        // A1 = 100 (integer constant in formula)
        wb.set_cell_formula("A1", "=100").await.unwrap();
        // B1 = 3.14 (float constant)
        wb.set_cell_formula("B1", "=3.14").await.unwrap();
        // C1 = TRUE
        wb.set_cell_formula("C1", "=TRUE()").await.unwrap();
        // D1 = -A1 (unary minus)
        wb.set_cell_formula("D1", "=-A1").await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

    assert_eq!(formula_text(&sheet.get_value_at(0, 0)), "=100");
    assert_eq!(formula_text(&sheet.get_value_at(0, 1)), "=3.14");
    // TRUE() in BIFF8 is stored as tFuncV for TRUE (idx 34) with 0 args
    let c1_val = sheet.get_value_at(0, 2);
    let c1_text = formula_text(&c1_val);
    assert!(
        c1_text == "=TRUE()" || c1_text == "=TRUE",
        "C1 should be =TRUE() or =TRUE, got {c1_text}"
    );
    assert_eq!(formula_text(&sheet.get_value_at(0, 3)), "=-A1");

    cleanup_fixture(&path);
}

#[test]
fn test_xls_formula_parentheses() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", 2.0).await.unwrap();
        wb.set_cell_value("B1", 3.0).await.unwrap();
        wb.set_cell_value("C1", 4.0).await.unwrap();
        // D1 = (A1+B1)*C1
        wb.set_cell_formula("D1", "=(A1+B1)*C1").await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

    assert_eq!(
        formula_text(&sheet.get_value_at(0, 3)),
        "=(A1+B1)*C1"
    );

    cleanup_fixture(&path);
}

#[test]
fn test_xls_formula_cached_values() {
    skip_if_no_lo!();
    let path = temp_fixture_path();

    runtime().block_on(async {
        let lo = lo_bridge().await.unwrap();
        let mut b = lo.lock().await;
        let mut wb = b.create_workbook().await.unwrap();
        wb.set_cell_value("A1", 10.0).await.unwrap();
        wb.set_cell_value("A2", 20.0).await.unwrap();
        // B1 = SUM(A1:A2) — cached numeric result should be 30
        wb.set_cell_formula("B1", "=SUM(A1:A2)").await.unwrap();
        wb.save_as_xls(path.to_str().unwrap()).await.unwrap();
        wb.close().await.unwrap();
    });

    let workbook = XlsReader::read_file(&path).unwrap();
    let sheet = workbook.worksheet(0).unwrap();

    let val = sheet.get_value_at(0, 1);
    match &val {
        CellValue::Formula {
            text,
            cached_value,
            ..
        } => {
            assert_eq!(text, "=SUM(A1:A2)", "formula text");
            match cached_value.as_deref() {
                Some(CellValue::Number(n)) => {
                    assert!(
                        (*n - 30.0).abs() < f64::EPSILON,
                        "cached value should be 30.0, got {n}"
                    );
                }
                other => panic!("Expected cached Number(30.0), got {:?}", other),
            }
        }
        other => panic!("Expected Formula, got {:?}", other),
    }

    cleanup_fixture(&path);
}
